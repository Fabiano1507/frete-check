from flask import Flask, render_template, request, send_file
import pandas as pd
import xml.etree.ElementTree as ET
import io
from datetime import datetime

app = Flask(__name__)

# =========================
# Funções auxiliares
# =========================

def ler_cte_xml(arquivo):
    tree = ET.parse(arquivo)
    root = tree.getroot()

    ns = {'cte': 'http://www.portalfiscal.inf.br/cte'}

    numero_cte = root.find('.//cte:nCT', ns)
    cidade_origem = root.find('.//cte:xMunIni', ns)
    cidade_destino = root.find('.//cte:xMunFim', ns)
    peso = root.find('.//cte:qCarga', ns)
    valor = root.find('.//cte:vTPrest', ns)

    return {
        'cte': numero_cte.text if numero_cte is not None else 'N/A',
        'origem': cidade_origem.text.upper().strip() if cidade_origem is not None else '',
        'destino': cidade_destino.text.upper().strip() if cidade_destino is not None else '',
        'peso': float(peso.text) if peso is not None else 0.0,
        'valor_cobrado': float(valor.text) if valor is not None else 0.0
    }


def buscar_valor_tabela(df, origem, destino, peso):
    linha = df[
        (df['origem'] == origem) &
        (df['destino'] == destino) &
        (df['peso_min'] <= peso) &
        (df['peso_max'] >= peso)
    ]

    if linha.empty:
        return 0.0

    return float(linha.iloc[0]['valor'])


# =========================
# Rotas
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

    # Padroniza texto da tabela
    df_tabela['origem'] = df_tabela['origem'].str.upper().str.strip()
    df_tabela['destino'] = df_tabela['destino'].str.upper().str.strip()

    resultados = []

    for xml in xml_files:
        dados = ler_cte_xml(xml)

        valor_tabela = buscar_valor_tabela(
            df_tabela,
            dados['origem'],
            dados['destino'],
            dados['peso']
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
            'peso': round(dados['peso'], 2),
            'valor_tabela': round(valor_tabela, 2),
            'valor_cobrado': round(dados['valor_cobrado'], 2),
            'diferenca': round(diferenca, 2),
            'status': status
        })

    app.config['ultimo_resultado'] = resultados

    return render_template('resultado.html', resultados=resultados)


@app.route('/exportar_excel')
def exportar_excel():

    resultados = app.config.get('ultimo_resultado')

    if not resultados:
        return "Nenhum resultado para exportar", 400

    df = pd.DataFrame(resultados)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Conferencia')

    output.seek(0)

    nome = f"conferencia_frete_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return send_file(
        output,
        download_name=nome,
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# =========================
# Inicialização
# =========================

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)