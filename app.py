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

    def get_text(path):
        el = root.find(path, ns)
        return el.text.strip() if el is not None and el.text else ''

    def get_float(path):
        el = root.find(path, ns)
        try:
            return float(el.text)
        except:
            return 0.0

    return {
        'cte': get_text('.//cte:nCT'),
        'origem': get_text('.//cte:xMunIni'),
        'destino': get_text('.//cte:xMunFim'),
        'peso': get_float('.//cte:qCarga'),
        'valor_cobrado': get_float('.//cte:vTPrest')
    }


def buscar_valor_tabela(df, origem, destino, peso):
    linha = df[
        (df['origem'].str.upper() == origem.upper()) &
        (df['destino'].str.upper() == destino.upper()) &
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
        return "Erro: selecione os arquivos corretamente", 400

    xml_files = request.files.getlist('xml')
    tabela_file = request.files['tabela']

    if not xml_files or tabela_file.filename == '':
        return "Erro: arquivos não selecionados", 400

    df_tabela = pd.read_excel(tabela_file)
    df_tabela.columns = [c.lower() for c in df_tabela.columns]

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
            'CT-e': dados['cte'],
            'Origem': dados['origem'],
            'Destino': dados['destino'],
            'Peso': round(dados['peso'], 4),
            'Valor Tabela': round(valor_tabela, 2),
            'Valor Cobrado': round(dados['valor_cobrado'], 2),
            'Diferença': round(diferenca, 2),
            'Status': status
        })

    app.config['ULTIMO_RESULTADO'] = resultados

    return render_template('resultado.html', resultados=resultados)


@app.route('/exportar_excel')
def exportar_excel():
    resultados = app.config.get('ULTIMO_RESULTADO')

    if not resultados:
        return "Nenhum resultado para exportar", 400

    df = pd.DataFrame(resultados)

    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        df.to_excel(writer, index=False, sheet_name='Conferência')

    output.seek(0)

    nome_arquivo = f"conferencia_frete_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return send_file(
        output,
        download_name=nome_arquivo,
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# =========================
# Inicialização
# =========================

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
