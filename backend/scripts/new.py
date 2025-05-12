import mysql.connector
import json
import sys
import pandas as pd
from datetime import datetime

# Recebe a data por argumento (YYYY-MM-DD)
data_filtro = sys.argv[1] if len(sys.argv) > 1 else None
if not data_filtro:
    print(json.dumps({"erro": "Data não fornecida"}))
    sys.exit()

try:
    conn = mysql.connector.connect(
        host="45.179.90.229",
        user="planejamento",
        password="899605aA@",
        database="vieira_online"
    )
    cursor = conn.cursor(dictionary=True)

    # Consulta contratos do dia
    query = """
        SELECT
            banco_refinanciador AS 'Banco',
            proposta_id_banco AS 'Nº Contrato',
            valor_parcela AS 'Valor Parcela',
            prazo AS 'Prazo',
            cliente_nome AS 'Nome',
            cliente_cpf AS 'CPF',
            convenio_nome AS 'Convênio',
            status_api AS 'Status (online)',
            matricula AS 'Matricula',
            valor_financiado AS 'Valor Bruto',
            valor_referencia AS 'Valor Referência',
            data_cadastro AS 'Data Cadastro',
            banco_nome AS 'Banco Originador',
            taxa AS 'taxa'
        FROM cadastrados
        WHERE DATE(data_cadastro) = %s
          AND empresa IN ('abbcred', 'franquiasazul', 'impacto')
    """
    cursor.execute(query, (data_filtro,))
    resultados = cursor.fetchall()

    if not resultados:
        print(json.dumps({"erro": "Nenhum contrato encontrado para a data"}))
        sys.exit()

    # Gerar planilha com os resultados
    # Carregar modelo da planilha original para manter a estrutura
    colunas_modelo = [
        'Banco', 'Nº Contrato', 'Valor Parcela', 'Prazo', 'Nome', 'CPF', 'CODIGO TABELA',
        'Convênio', 'Status (online)', 'Último Comentário (online)', 'CODIGO DE LOJA',
        'Matricula', 'Valor Bruto', 'Valor Referência', 'Data Cadastro',
        'Banco Originador', 'Data Retorno Cip', 'taxa'
    ]

    # Preencher com os dados disponíveis
    df_resultado = pd.DataFrame(resultados)

    # Criar DataFrame final com todas as colunas do modelo (as não existentes ficam em branco)
    df_final = pd.DataFrame(columns=colunas_modelo)
    for col in df_resultado.columns:
        if col in df_final.columns:
            df_final[col] = df_resultado[col]

    # Salvar a planilha
    nome_arquivo = f"relatorio_{data_filtro}.xlsx"
    caminho = f"./{nome_arquivo}"
    df_final.to_excel(caminho, index=False)

    print(json.dumps({"arquivo": nome_arquivo}))

except Exception as e:
    print(json.dumps({"erro": str(e)}))
finally:
    if conn.is_connected():
        cursor.close()
        conn.close()
