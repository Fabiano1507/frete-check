print(">>> APP.PY CARREGADO <<<")

import os
import math
import pandas as pd
import xml.etree.ElementTree as ET
from flask import Flask, render_template, request, send_file
from io import BytesIO

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# -------------------------
# CONFIGURAÇÃO POR CLIENTE
# -------------------------
CLIENTES = {
    "britania_joinville": {
        "frete": "britania_joinville.csv",
        "icms": "tabela_icms_dentro_origem_sc.csv",
        "origem_nome": "Joinville",
        "uf_origem": "SC"
    },
    "britania_linhares": {
        "frete": "britania_linhares.csv",
        "icms": "tabela_icms_dentro_origem_es.csv",
        "origem_nome": "Linhares",
        "uf_origem": "ES"
    }
}

# Tabela fixa
TABELA_REGIOES = pd.read_csv(
    os.path.join(BASE_DIR, "tabelas", "regioes_grande_capital.csv")
)

# -------------------------
# UTILIDADES
# -------------------------
def identificar_tipo_destino(uf, cidade):
    cidade = cidade.upper().strip()
    linha = TABELA_REGIOES[
        (TABELA_REGIOES["uf"] == uf) &
        (TABELA_REGIOES["cidade"] == cidade)
    ]
    return linha.iloc[0]["regiao"] if not linha.empty else "INTERIOR"


def obter_divisor_icms(tabela_icms, uf_origem, uf_destino):
    linha = tabela_icms[
        (tabela_icms["uf_origem"] == uf_origem) &
        (tabela_icms["uf_destino"] == uf_destino)
    ]
    return linha.iloc[0]["divisor"] if not linha.empty else 1


# -------------------------
# LEITURA XML CT-e
# -------------------------
def ler_cte_xml(arquivo, origem_nome, uf_origem):
    ns = {"ns": "http://www.portalfiscal.inf.br/cte"}
    root = ET.parse(arquivo).getroot()

    def get_text(path):
        el = root.find(path, ns)
        return el.text.strip() if el is not None else ""

    cte = {}
    cte["numero"] = get_text(".//ns:nCT")
    cte["uf_origem"] = uf_origem
    cte["uf_destino"] = get_text(".//ns:UFFim")
    cte["cidade_destino"] = get_text(".//ns:xMunFim")
    cte["origem"] = origem_nome

    peso_declarado = 0
    peso_cubado = 0
    m3 = 0

    for infq in root.findall(".//ns:infQ", ns):
        tp = infq.find("ns:tpMed", ns)
        if tp is None:
            continue

        valor = float(infq.find("ns:qCarga", ns).text.replace(",", "."))

        if tp.text.upper() == "PESO DECLARADO":
            peso_declarado = valor
        elif tp.text.upper() == "PESO BASE DE CALCULO":
            peso_cubado = valor
        elif tp.text.upper() == "PESO CUBADO":
            m3 = valor

    cte["peso_declarado"] = peso_declarado
    cte["peso_cubado"] = peso_cubado
    cte["m3"] = m3

    valor_cte = root.find(".//ns:vTPrest", ns)
    cte["valor_cobrado"] = float(valor_cte.text.replace(",", "."))

    v_carga = root.find(".//ns:vCarga", ns)
    cte["valor_mercadoria"] = float(v_carga.text.replace(",", ".")) if v_carga is not None else 0

    return cte


# -------------------------
# CÁLCULO DO FRETE
# -------------------------
def calcular_frete(cte, tabela_frete, tabela_icms):
    tipo_destino = identificar_tipo_destino(cte["uf_destino"], cte["cidade_destino"])

    linha = tabela_frete[
        (tabela_frete["uf_destino"] == cte["uf_destino"]) &
        (tabela_frete["tipo_destino"] == tipo_destino)
    ]

    if linha.empty:
        return None

    l = linha.iloc[0]

    frete_m3 = cte["m3"] * l["valor_m3"]
    frete_base = max(frete_m3, l["frete_min"])

    adv = cte["valor_mercadoria"] * l["adv_valor"]
    taxa = l["taxa"]

    faixas = math.ceil(cte["peso_declarado"] / 100)
    pedagio = faixas * l["pedagio"]

    subtotal = frete_base + adv + taxa + pedagio

    divisor = obter_divisor_icms(
        tabela_icms,
        cte["uf_origem"],
        cte["uf_destino"]
    )

    valor_tabela = round(subtotal / divisor, 2)
    valor_cobrado = round(cte["valor_cobrado"], 2)
    diferenca = round(valor_cobrado - valor_tabela, 2)

    if abs(diferenca) <= 0.01:
        status = "OK"
    elif diferenca > 0:
        status = "MAIOR"
    else:
        status = "MENOR"

    return {
        "cte": cte["numero"],
        "origem": cte["origem"],
        "destino": f"{cte['cidade_destino']}/{cte['uf_destino']}",
        "valor_tabela": valor_tabela,
        "valor_cobrado": valor_cobrado,
        "diferenca": diferenca,
        "status": status,
        "memoria": {
            "frete_m3": frete_m3,
            "frete_base": frete_base,
            "adv": adv,
            "taxa": taxa,
            "pedagio": pedagio,
            "divisor_icms": divisor
        }
    }


# -------------------------
# ROTAS
# -------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/processar", methods=["POST"])
def processar():
    cliente = request.form.get("cliente")
    config = CLIENTES.get(cliente)

    tabela_frete = pd.read_csv(os.path.join(BASE_DIR, "tabelas", config["frete"]))
    tabela_icms = pd.read_csv(os.path.join(BASE_DIR, "tabelas", config["icms"]))

    arquivos = request.files.getlist("xmls")
    resultados = []

    for arq in arquivos:
        cte = ler_cte_xml(
            arq,
            config["origem_nome"],
            config["uf_origem"]
        )
        res = calcular_frete(cte, tabela_frete, tabela_icms)
        if res:
            resultados.append(res)

    df = pd.DataFrame(resultados)

    totais = {
        "total_tabela": round(df["valor_tabela"].sum(), 2),
        "total_cobrado": round(df["valor_cobrado"].sum(), 2),
        "total_diferenca": round(df["diferenca"].sum(), 2)
    }

    app.config["ULTIMO_RESULTADO"] = resultados

    return render_template(
        "resultado.html",
        resultados=resultados,
        totais=totais
    )


@app.route("/exportar", methods=["POST"])
def exportar():
    resultados = app.config.get("ULTIMO_RESULTADO", [])

    df = pd.DataFrame([{
        "CTe": r["cte"],
        "Origem": r["origem"],
        "Destino": r["destino"],
        "Valor Tabela": r["valor_tabela"],
        "Valor Cobrado": r["valor_cobrado"],
        "Diferença": r["diferenca"],
        "Status": r["status"]
    } for r in resultados])

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="resultado_frete.xlsx",
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
