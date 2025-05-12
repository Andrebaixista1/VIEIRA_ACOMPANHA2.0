import json, sys, time, pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoAlertPresentException, TimeoutException, StaleElementReferenceException
)

def log(msg):
    sys.stderr.write(f"[LOG] {msg}\n")

# Mapeamento de colunas
mapeamento = {
    "PROPOSTA": "Nº Contrato",
    "CPF": "CPF",
    "CLIENTE": "Nome",
    "CONVENIO": "Convênio",
    "ATIVIDADE": "Status (online)",
    "VLR PARCELA": "Valor Parcela",
    "VALOR SOLICITADO": "Valor Referência"
}
colunas_modelo = [
    'Banco','Nº Contrato','Valor Parcela','Prazo','Nome','CPF','CODIGO TABELA',
    'Convênio','Status (online)','Último Comentário (online)','CODIGO DE LOJA',
    'Matricula','Valor Bruto','Valor Referência','Data Cadastro',
    'Banco Originador','Data Retorno Cip','taxa'
]

# Data passada por argumento
data_filtro = sys.argv[1] if len(sys.argv)>1 else None
if not data_filtro:
    print(json.dumps({"erro":"Data não fornecida"}))
    sys.exit()

try:
    # configurações do Chrome headless
    log("Configurando Chrome headless.")
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=opts)

    # 1) LOGIN
    log("Acessando login do C6.")
    driver.get("https://c6.c6consig.com.br/WebAutorizador/Login/AC.UI.LOGIN.aspx?FISession=a2896b98e8b")
    time.sleep(2)
    log("Preenchendo credenciais.")
    driver.find_element(By.ID,"EUsuario_CAMPO").send_keys("52740689870_000435")
    driver.find_element(By.ID,"ESenha_CAMPO").send_keys("Brabox@25")
    driver.find_element(By.ID,"lnkEntrar").click()
    time.sleep(2)
    try:
        alert = driver.switch_to.alert
        log("Confirmando alerta.")
        alert.accept()
        time.sleep(1)
    except NoAlertPresentException:
        log("Nenhum alerta.")

    # 2) ABRIR DROPDOWN Esteira
    log("Clicando no toggle 'Esteira'.")
    toggle = WebDriverWait(driver,30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR,"a.dropdown-toggle[data-submenu]"))
    )
    driver.execute_script("arguments[0].click();",toggle)
    time.sleep(1)

    # 3) CLICAR EM Aprovação / Consulta via href
    log("Aguardando dropdown-menu e clicando no link 'Aprovação / Consulta'.")
    WebDriverWait(driver,30).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR,"ul.dropdown-menu"))
    )
    link = driver.find_element(
        By.CSS_SELECTOR,
        "ul.dropdown-menu li a[href*='AprovacaoConsulta']"
    )
    driver.execute_script("arguments[0].click();",link)
    time.sleep(5)

    # 4) TELA DE FILTROS
    log("Aguardando tela de filtros.")
    WebDriverWait(driver,30).until(
        EC.presence_of_element_located((By.ID,"ctl00_Cph_AprCons_cbxPesquisaPor_CAMPO"))
    )
    log("Selecionando 'Data Base'.")
    Select(driver.find_element(By.ID,"ctl00_Cph_AprCons_cbxPesquisaPor_CAMPO"))\
        .select_by_visible_text("Data Base")

    log(f"Preenchendo data: {data_filtro}")
    # tenta até 3x em caso de stale
    for attempt in range(3):
        try:
            fld = WebDriverWait(driver,30).until(
                EC.presence_of_element_located((By.ID,"ctl00_Cph_AprCons_txtPesquisa_CAMPO"))
            )
            fld.clear()
            fld.send_keys(data_filtro)
            break
        except StaleElementReferenceException:
            log(f"StaleElementReference ao preencher data, tentativa {attempt+1}/3.")
            time.sleep(1)
    else:
        raise Exception("Não foi possível preencher o campo de data devido a stale element.")

    log("Clicando em 'Pesquisar'.")
    WebDriverWait(driver,30).until(
        EC.element_to_be_clickable((By.ID,"btnPesquisar_txt"))
    ).click()
    time.sleep(5)

    # 5) FUNÇÕES DE SCRAPING
    def extrair_cabecalho():
        log("Tentando extrair cabeçalho da tabela.")
        try:
            hdr = WebDriverWait(driver,10).until(
                EC.presence_of_element_located((By.XPATH,
                    "//table[@id='ctl00_Cph_AprCons_grdConsulta']//tr[@class='header']"
                ))
            )
            cols = [th.text.strip().upper() for th in hdr.find_elements(By.TAG_NAME,"th")]
            log(f"Colunas detectadas: {cols}")
            return cols
        except TimeoutException:
            log("Cabeçalho não encontrado: nenhum resultado.")
            return []

    def extrair_linhas(headers):
        if not headers:
            return []
        log("Extraindo linhas da tabela.")
        rows = driver.find_elements(By.XPATH,
            "//table[@id='ctl00_Cph_AprCons_grdConsulta']//tr[@class='normal' or @class='alternate']"
        )
        result = []
        for r in rows:
            cells = r.find_elements(By.TAG_NAME,"td")
            vals = [c.text.strip() for c in cells]
            rec = {c:"" for c in colunas_modelo}
            rec['Banco'] = 'C6'
            rec['Data Cadastro'] = data_filtro
            for i,v in enumerate(vals):
                if i < len(headers) and headers[i] in mapeamento:
                    rec[mapeamento[headers[i]]] = v
            result.append(rec)
        log(f"{len(result)} linhas extraídas.")
        return result

    # 6) EXECUTA SCRAPING
    dados = []
    headers = extrair_cabecalho()
    page = 1
    if headers:
        while True:
            log(f"Processando página {page}.")
            dados.extend(extrair_linhas(headers))
            try:
                nxt = WebDriverWait(driver,10).until(
                    EC.element_to_be_clickable((By.ID,"ctl00_Cph_AprCons_lkbProximo"))
                )
                nxt.click()
                time.sleep(5)
                page += 1
            except TimeoutException:
                log("Sem próxima página.")
                break

    driver.quit()

    # 7) GERAR PLANILHA (mesmo que vazia)
    log("Gerando planilha.")
    df = pd.DataFrame(dados, columns=colunas_modelo)
    arquivo = f"relatorio_c6_{data_filtro}.xlsx"
    df.to_excel(arquivo, index=False)
    log(f"Planilha criada: {arquivo}")
    print(json.dumps({"arquivo":arquivo}))

except Exception as e:
    import traceback
    log("Erro fatal:\n"+traceback.format_exc())
    try:
        driver.save_screenshot(f"erro_{data_filtro}.png")
        log("Screenshot salvo.")
    except:
        log("Falha ao salvar screenshot.")
    print(json.dumps({"erro":str(e)}))
    sys.exit()
