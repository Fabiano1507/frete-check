from flask import Flask, render_template, request, send_file
import pandas as pd
import xml.etree.ElementTree as ET
import io
from datetime import datetime
import os

app = Flask(__name__)

# =========================
# CONFIGURAÇÕES
# =========================

FATOR_CUBAGEM = 300  # kg por m³ (ajustável por cliente)

# =========================
# FUNÇÕES AUXILIARES
# =========================

def ler_cte_xml(arquivo):
    """
    Lê os principais dados do CT-e:
    - Número do CT-e
    - UF origem
    - UF destino
    - Peso cubado (kg)
    - Valor cobrado
    """
    tree = ET.parse(arquivo)
    root = tree.getroot()

    ns = {'cte': 'http://www.portalfiscal.inf.br/cte'}

    numero_cte = root.find('.//cte:nCT', ns)
    uf_origem = root.find('.//cte:UFIni', ns)
    uf_destino = root.find('.//cte:UFFim', ns)
    peso = root.find('.//cte:qCarga', ns)
    valor = root.find('.//cte:vTPrest', ns)

    peso_cubado = float(peso.text) if peso is not None else 0.0
    valor_cobrado = float(valor.text) if valor is not None else 0.0

    m3 = round(peso_cubado / FATOR_CUBAGEM, 4) if peso_cubado > 0 else 0.0

    return {
        'cte': numero_cte.text if numero_cte is not None else 'N/A',
        'origem': uf_origem.text.upper() if uf_origem is not None else '',
        'destino': uf_destino.text.upper() if uf_destino is not None else '',
        'peso_cubado': peso_cubado,
        'm3': m3,
        'valor_cobrado': valor_cobrado
    }


def identificar_regiao(destino):
    """
    Regra simples inicial:
    Pode ser expandida depois (capital/interior por UF ou cidade)
    """
    capitais = [
        'SP', 'RJ', 'BH', 'POA', 'CURITIBA', 'SALVADOR', 'RECIFE',
        'FORTALEZA', 'BELO HORIZONTE'
    ]

    if destino in ['SP', 'RJ', 'MG', 'BA', 'CE', 'PE', 'RS', 'PR']:
        return 'CAPITAL'
    else:
        return 'INTERIOR'


def buscar_valor_tabela(df, percurso, regiao, m3):
    """
    Busca o valor correto na tabela Britânia (ou similar)
    """
    linha = df[
        (df['percurso'] == percurso) &
        (df['regiao'] == regiao) &
        (df['m3_min'] <= m3) &
        (df['m3_max'] >= m3)
    ]

    if linha.empty:
        return 0.0

    return float(linha.iloc[0]['valor'])


# =========================
# ROTAS
# =========================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/processar', methods=['POST'])
def processar():
    if 'xml' not in request.files or 'tabela' not in request.files:
        return "Erro: selecione os XMLs e a tabela.", 400

    xml_files = request.files.getlist('xml')
    tabela_file = request.files['tabela']

    if not xml_files or tabela_file.filename == '':
        return "Erro: arquivos não selecionados.", 400

    df_tabela = pd.read_excel(tabela_file)
    df_tabela.columns = [c.lower() for c in df_tabela.columns]

    resultados = []

    for xml_file in xml_files:
        dados = ler_cte_xml(xml_file)

        percurso = f"{dados['origem']} x {dados['destino']}"
        regiao = identificar_regiao(dados['destino'])

        valor_tabela = buscar_valor_tabela(
            df_tabela,
            percurso,
            regiao,
            dados['m3']
        )

        diferenca = dados['valor_cobrado'] - valor_tabela

        if diferenca > 0:
            status = 'A MAIOR'
        elif diferenca < 0:
            status = 'A MENOR'
        else:
            status = 'OK'

        resultados.append({
            'cte': dados['cte'],
            'origem': dados['origem'],
            'destino': dados['destino'],
            'peso_cubado': dados['peso_cubado'],
            'm3': dados['m3'],
            'valor_tabela': valor_tabela,
            'valor_cobrado': dados['valor_cobrado'],
            'diferenca': diferenca,
            'status': status
        })

    app.config['ultimo_resultado'] = resultados

    return render_template('resultado.html', resultados=resultados)


@app.route('/exportar_excel')
def exportar_excel():
    resultados = app.config.get('ultimo_resultado')

    if not resultados:
        return "Nenhum resultado para exportar.", 400

    df = pd.DataFrame(resultados)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Conferencia')

    output.seek(0)

    nome_arquivo = f"conferencia_frete_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# =========================
# START
# =========================

if __name__ == '__main__':
    app.run(debug=True)
