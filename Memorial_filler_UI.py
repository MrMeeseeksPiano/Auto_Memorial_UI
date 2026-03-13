import streamlit as st
import pdfplumber
import re
import math
import io
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
import sqlite3

rotulos = [
    "CEP",
    "Nome",
    "Endereco",
    "Telefone",
    "Email",
    "Qtd_mod",
    "Qtd_inv",
    "Pot_nom",
]
rotulos_pdf = [
    "CEP da UC com GD",
    "Nome do Titular da UC com GD",
    "Endereço",
    "Telefone do Titular \(DDD \+ número\)",
    "E-mail do Titular da UC com GD",
    "Quantidade de Módulos",
    "Quantidade de Inversores",
    "Potência Total dos Módulos \(kW\)",
]
energia_mensal = {
    "Energia_jan": 0,
    "Energia_fev": 0,
    "Energia_mar": 0,
    "Energia_abr": 0,
    "Energia_maio": 0,
    "Energia_jun": 0,
    "Energia_jul": 0,
    "Energia_ago": 0,
    "Energia_set": 0,
    "Energia_out": 0,
    "Energia_nov": 0,
    "Energia_dez": 0,
}
irradiacao = [5.01, 5.5, 5.1, 5.46, 5.56, 5.61, 5.83, 6.47, 5.91, 5.45, 4.75, 4.98]
mes = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

# Interface
st.title("Gerador de Memorial Descritivo")

conn = sqlite3.connect("equipamentos.db")
cursor = conn.cursor()

cursor.execute("SELECT id, Fabricante, Pot_max FROM Modulos ")
conteudo_modulos = cursor.fetchall()
opcoes_mod = []
mapeamento_mod = {}

cursor.execute("SELECT id, Fabricante_sigla FROM Inversores")
conteudo_inversores = cursor.fetchall()
opcoes_inv = []
mapeamento_inv = {}

for id_mod,fabricante,potencia,in conteudo_modulos:
    lista_mod = f"{fabricante} {potencia}"
    opcoes_mod.append(lista_mod)
    mapeamento_mod[lista_mod] = id_mod

for id_inv, inv in conteudo_inversores:
    lista_inv = f"{inv}"
    opcoes_inv.append(lista_inv)
    mapeamento_inv[lista_inv] = id_inv

conn.close()

mod_sel = st.selectbox("Módulo:", opcoes_mod)
id_gerador_escolhido = mapeamento_mod[mod_sel]

inv_sel = st.selectbox("Inversor:", opcoes_inv)
id_inversor_escolhido = mapeamento_inv[inv_sel]


opcoes_estrutura = ["Parafuso Prisioneiro", "Laje", "Solo", "Metálico"]
Estrutura_sel = st.selectbox("Tipo de estrutura:", opcoes_estrutura)
tipo_estrutura = opcoes_estrutura.index(Estrutura_sel)

opcoes_fornecimento = ["220 V", "380 V"]
fornecimento_sel = st.selectbox("Tipo de fornecimento:", opcoes_fornecimento)
tipo_fornecimento = opcoes_fornecimento.index(fornecimento_sel)


arquivo_pdf = st.file_uploader("Upload do PDF:", type="pdf")

if arquivo_pdf:
    with pdfplumber.open(arquivo_pdf) as pdf:
        first_page = pdf.pages[0]
        pdf_text = first_page.extract_text()

    def valor(label, texto):
        padrao = rf"{label}:\s*(.+)"
        encontrado = re.search(padrao, texto, re.IGNORECASE)
        return encontrado.group(1).strip()

    def valor_coordenada(label, texto):
        padrao = rf"{label}\s*(.+)"
        encontrado = re.search(padrao, texto, re.IGNORECASE | re.MULTILINE)
        return encontrado.group(1).strip()

    latitude = valor_coordenada("latitude", pdf_text)
    longitude = valor_coordenada("longitude", pdf_text)

    def gms_para_decimal(coordenada_gms_string):
        padrao = r"(\D)\s*(\d+)\s+(\d+)\s+([\d.,]+)"
        match = re.search(padrao, coordenada_gms_string, re.IGNORECASE)
        if not match:
            return "N/A"
        direcao = match.group(1).upper()
        grau = int(match.group(2))
        minuto = int(match.group(3))
        segundo = float(match.group(4).replace(",", "."))
        valor_decimal = grau + (minuto / 60) + (segundo / 3600)
        if direcao in ["S", "O", "W"]:
            valor_decimal *= -1
        return round(valor_decimal, 6)

    latitude_decimal = gms_para_decimal(latitude)
    longitude_decimal = gms_para_decimal(longitude)

    dicionario = {}
    i = 0
    for rotulo_item in rotulos:
        dicionario[rotulo_item] = valor(rotulos_pdf[i], pdf_text)
        i += 1

    Pot_ano = float(dicionario["Pot_nom"].replace(",", ".")) * 128 * 12
    Pot_ano_rounded = int(Pot_ano // 1)
    Pot_ano_rounded = Pot_ano_rounded - (Pot_ano_rounded % 100)
    dicionario["Pot_ano_rounded"] = str(Pot_ano_rounded)

    dicionario["Area_mod"] = valor("Área Total dos Arranjos \(m²\)", pdf_text)
    Qtd_mod_var = int(dicionario["Qtd_mod"])
    Area_mod_total = round(
        float(dicionario["Area_mod"].replace(",", ".")) * Qtd_mod_var, 2
    )
    dicionario["Area_mod_total"] = Area_mod_total

    dicionario["Latitude"] = latitude_decimal
    dicionario["Longitude"] = longitude_decimal

    Bairro_partes = dicionario["Endereco"].rsplit(",", 1)
    Bairro = Bairro_partes[1].strip()
    dicionario["Bairro"] = Bairro

    Pot_mensal = int((Pot_ano_rounded / 12) // 1)
    dicionario["Pot_mensal"] = Pot_mensal
    Pot_diaria = round(float(Pot_mensal / 30), 2)
    dicionario["Pot_diaria"] = Pot_diaria
    Pot_nom_calc = round(Pot_diaria / (0.8 * 5.34), 2)
    dicionario["Pot_nom_calc"] = Pot_nom_calc

    Nomes = dicionario["Nome"].split()
    Nome_capa = f"{Nomes[0]} {Nomes[-1]}"
    dicionario["Nome_capa"] = Nome_capa
    Nome_login = f"{Nomes[0]}{Nomes[-1]}"
    dicionario["Nome_login"] = Nome_login
    Senha_login = f"{Nomes[0][0]}{Nomes[-1][0]}123456".lower()
    dicionario["Senha_login"] = Senha_login

    # Tipo de estrutura
    dicionario["Estrutura_tipo"] = opcoes_estrutura[tipo_estrutura]

    # Tipo de fornecimento
    dicionario["Tipo_fornecimento"] = opcoes_fornecimento[tipo_fornecimento]

    # Escolha do gerador e módulo
    conn1 = sqlite3.connect("equipamentos.db")
    conn1.row_factory = sqlite3.Row
    cursor1 = conn1.cursor()
    cursor1.execute("SELECT * FROM Inversores WHERE id = ? ", (id_inversor_escolhido,))
    resultado_inv = cursor1.fetchone()
    dicionario.update(dict(resultado_inv))
  

    cursor1.execute("SELECT * FROM Modulos WHERE id = ? ", (id_gerador_escolhido,))
    resultado_mod = cursor1.fetchone()
    dicionario.update(dict(resultado_mod))
    conn1.close()

    Pot_nom_com_virgula = float(dicionario["Pot_nom"].replace(",", "."))
    Pot_max_lista = resultado_mod["Pot_max"].rsplit(" ", 1)
    Pot_max_valor = float(Pot_max_lista[0]) / 1000
    dicionario["Pot_max_valor"] = Pot_max_valor
    dicionario["N_mod"] = round((Pot_nom_calc / Pot_max_valor), 2)

    m = 0
    total_energia = 0
    for key in energia_mensal.keys():
        energia_mensal[key] = round(
            Pot_nom_com_virgula * mes[m] * irradiacao[m] * 0.8, 2
        )
        total_energia += energia_mensal[key]
        m += 1

    dicionario.update(energia_mensal)
    dicionario["Total_energia"] = round(total_energia, 2)
    dicionario["Total_arredondado"] = 10 * math.ceil(total_energia / 10)

    if st.button("Gerar Memorial Word"):
        try:
            doc = DocxTemplate("Memorial Descritivo - Template.docx")
            dicionario["imagem_gerador"] = InlineImage(
                doc, resultado_mod['Imagem_gerador'], width=Cm(4.0)
            )
            dicionario["imagem_inversor"] = InlineImage(
                doc, resultado_inv['Imagem_inversor'], width=Cm(3.5)
            )
            doc.render(dicionario)

            output_stream = io.BytesIO()
            doc.save(output_stream)
            output_stream.seek(0)

            st.download_button(
                label="Baixar Memorial",
                data=output_stream,
                file_name=f"Memorial Descritivo - {Nome_capa}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as e:
            st.error(f"Erro: {e}")
