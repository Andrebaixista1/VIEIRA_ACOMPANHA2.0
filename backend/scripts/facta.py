import json
import sys
import time
import os
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException

def log(msg):
    sys.stderr.write(f"[LOG] {msg}\n")

date_filter = sys.argv[1] if len(sys.argv) > 1 else None
if not date_filter:
    print(json.dumps({"erro": "Data obrigatória não fornecida"}))
    sys.exit()

try:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)

    log("Acessando login...")
    driver.get("https://desenv.facta.com.br/sistemaNovo/login.php")
    WebDriverWait(driver, 15).until(lambda d: d.execute_script("return document.readyState") == "complete")

    driver.find_element(By.ID, "login").send_keys("97248")
    driver.find_element(By.ID, "senha").send_keys("Expande68@")
    driver.find_element(By.ID, "btnLogin").click()
    time.sleep(2)

    try:
        driver.switch_to.alert.accept()
        log("Alerta aceito.")
    except NoAlertPresentException:
        log("Nenhum alerta pós-login.")

    driver.get("https://desenv.facta.com.br/sistemaNovo/andamentoPropostas.php")
    WebDriverWait(driver, 15).until(lambda d: d.execute_script("return document.readyState") == "complete")
    log(f"Página de propostas carregada. URL atual: {driver.current_url}")
    log(f"Título da página: {driver.title}")

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "periodoini")))
        log("Campo 'periodoini' encontrado.")
    except:
        driver.save_screenshot("erro_tela_periodoini.png")
        raise Exception("Página carregada, mas campo 'periodoini' não visível ou login falhou.")

    driver.find_element(By.ID, "periodoini").send_keys(date_filter)
    driver.find_element(By.ID, "periodofim").send_keys(date_filter)
    driver.find_element(By.CSS_SELECTOR, "input[name='tipoPeriodo'][value='data_status']").click()
    driver.find_element(By.ID, "pesquisar").click()
    time.sleep(2)

    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "tblListaProposta")))
    tabela = driver.find_element(By.ID, "tblListaProposta")

    cabecalho = tabela.find_elements(By.CSS_SELECTOR, "thead th")
    mapa = {
        "Banco": None,
        "Nº Contrato": None,
        "CPF": None,
        "Nome": None,
        "Convênio": None,
        "Valor Bruto": None,
        "Valor Parcela": None,
        "Prazo": None,
        "Status (online)": None
    }

    for i, th in enumerate(cabecalho):
        texto = th.text.strip().lower()
        if "banco" in texto: mapa["Banco"] = i
        elif "contrato" in texto: mapa["Nº Contrato"] = i
        elif "cpf" in texto: mapa["CPF"] = i
        elif "cliente" in texto or "nome" in texto: mapa["Nome"] = i
        elif "averbador" in texto: mapa["Convênio"] = i
        elif "valor" in texto: mapa["Valor Bruto"] = i
        elif "parcela" in texto: mapa["Valor Parcela"] = i
        elif "prazo" in texto: mapa["Prazo"] = i
        elif "status" in texto: mapa["Status (online)"] = i

    dados = []

    def extrair_linhas():
        linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            if not colunas or len(colunas) < 5: continue
            bruto = colunas[mapa["Valor Bruto"]].text.strip() if mapa["Valor Bruto"] is not None else ""
            dados.append({
                "Banco": colunas[mapa["Banco"]].text.strip() if mapa["Banco"] is not None else "",
                "Nº Contrato": colunas[mapa["Nº Contrato"]].text.strip() if mapa["Nº Contrato"] is not None else "",
                "CPF": colunas[mapa["CPF"]].text.strip() if mapa["CPF"] is not None else "",
                "Nome": colunas[mapa["Nome"]].text.strip() if mapa["Nome"] is not None else "",
                "Convênio": colunas[mapa["Convênio"]].text.strip() if mapa["Convênio"] is not None else "",
                "Valor Bruto": bruto,
                "Valor Referência": bruto,
                "Valor Parcela": colunas[mapa["Valor Parcela"]].text.strip() if mapa["Valor Parcela"] is not None else "",
                "Prazo": colunas[mapa["Prazo"]].text.strip() if mapa["Prazo"] is not None else "",
                "Status (online)": colunas[mapa["Status (online)"]].text.strip() if mapa["Status (online)"] is not None else ""
            })

    # Paginação
    while True:
        extrair_linhas()
        try:
            proximo = driver.find_element(By.ID, "paginacaoProximo")
            if not proximo.is_enabled() or "disabled" in proximo.get_attribute("class").lower():
                log("Botão 'Próximo' desabilitado.")
                break
            log("Clicando em 'Próximo'...")
            proximo.click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "tblListaProposta")))
            tabela = driver.find_element(By.ID, "tblListaProposta")
        except:
            log("Fim da paginação ou botão 'Próximo' não disponível.")
            break

    if not dados:
        raise Exception("Tabela carregada, mas nenhum dado extraído.")

    df = pd.DataFrame(dados, columns=[
        "Banco", "Nº Contrato", "CPF", "Nome", "Convênio",
        "Valor Bruto", "Valor Referência", "Valor Parcela", "Prazo", "Status (online)"
    ])

    # Define pasta Documentos\relatorios
    documents_path = os.path.join(os.path.expanduser("~"), "Documents", "relatorios")
    os.makedirs(documents_path, exist_ok=True)
    nome_arquivo = os.path.join(documents_path, f"facta_{date_filter}.xlsx")

    df.to_excel(nome_arquivo, index=False)
    log(f"Planilha gerada: {nome_arquivo}")
    print(json.dumps({"arquivo": nome_arquivo}))
    driver.quit()

except Exception as e:
    log("Erro fatal:")
    log(traceback.format_exc())
    print(json.dumps({"erro": str(e)}))
    sys.exit()
