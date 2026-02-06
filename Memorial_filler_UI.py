import streamlit as st
import pdfplumber
import re
import math
import io
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

rotulos = ['CEP','Nome','Endereco','Telefone','Email','Qtd_mod','Qtd_inv','Pot_nom']
rotulos_pdf = ['CEP da UC com GD','Nome do Titular da UC com GD','Endereço','Telefone do Titular \(DDD \+ número\)','E-mail do Titular da UC com GD','Quantidade de Módulos','Quantidade de Inversores','Potência Total dos Módulos \(kW\)']
energia_mensal = {'Energia_jan':0,'Energia_fev':0,'Energia_mar':0,'Energia_abr':0,'Energia_maio':0,'Energia_jun':0,'Energia_jul':0,'Energia_ago':0,'Energia_set':0,'Energia_out':0,'Energia_nov':0,'Energia_dez':0}
irradiacao = [5.01, 5.5, 5.1, 5.46, 5.56, 5.61, 5.83, 6.47, 5.91, 5.45, 4.75, 4.98]
mes = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

gerador_trina = {'Fabricante':'Trina Solar','SIGLA':'TSM-695NEG21C.20','Tec_construcao':'Monocristalino','Garantia':'12 anos','Pot_max':'695 W','Eficiencia':'22,4 %','Tensao_nom':'40,3 V','Tensao_aberto':'48,3 V','Corrente_nom':'17,25 A','Corrente_cc':'18,28 A','axlxp':'2384 x 1303 x 33 mm','Peso':'38,3 kg','Imagem_gerador':'Imagens/Trina_gerador.jpg'}
gerador_astronergy = {'Fabricante': 'Astronergy','SIGLA': 'CHSM6612M/HV - 375W','Tec_construcao': 'Monocristalino','Garantia': '10 anos','Pot_max': '375 W','Eficiencia': '19,4 %','Tensao_nom': '39,76 V','Tensao_aberto': '48,45 V','Corrente_nom': '9,45 A','Corrente_cc': '9,94 A','axlxp': '1960 x 992 x 40 mm','Peso': '21,8 kg','Imagem_gerador':'Imagens/Trina_gerador.jpg'}
gerador_astronergy_600 = {'Fabricante': 'Astronergy','SIGLA': 'CHSM66RN(DG)/F-BH-600W','Tec_construcao': 'Monocristalino','Garantia': '12 anos','Pot_max': '600 W','Eficiencia': '22,2 %','Tensao_nom': '41,05 V','Tensao_aberto': '48,44 V','Corrente_nom': '14,64 A','Corrente_cc': '15,78 A','axlxp': '2382 x 1134 x 30 mm','Peso': '33,5 kg','Imagem_gerador': 'Imagens/Trina_gerador.jpg'}
gerador_canadian = {'Fabricante':'Canadian Solar','SIGLA':'TCS6U-330P','Tec_construcao':'Policristalino','Garantia':'12 anos','Pot_max':'330 W','Eficiencia':'16,97 %','Tensao_nom':'37,2 V','Tensao_aberto':'45,6 V','Corrente_nom':'8,88 A','Corrente_cc':'9,45 A','axlxp':'1990 x 992 x 40 mm','Peso':'22,4 kg','Imagem_gerador':'Imagens/Trina_gerador.jpg'}
gerador_gokin = {'Fabricante': 'Gokin Solar','SIGLA': 'GK-1-72HT585M','Tec_construcao': 'Monocristalino','Garantia': '12 anos','Pot_max': '585 W','Eficiencia': '22,6%','Tensao_nom': '42.74 V','Tensao_aberto': '51.67 V','Corrente_nom': '13.69 A','Corrente_cc': '14.43 A','axlxp': '2310x1125x1259mm','Peso': '26,8 kg','Imagem_gerador':'Imagens/Trina_gerador.jpg'}
gerador_gokin_700 = {'Fabricante': 'Gokin Solar','SIGLA': 'GK-2-66HTBD-700M','Tec_construcao': 'Monocristalino','Garantia': '12 anos','Pot_max': '700 W','Eficiencia': '22,5 %','Tensao_nom': '41,1 V','Tensao_aberto': '47,9 V','Corrente_nom': '17,04 A','Corrente_cc': '18,8 A','axlxp': '2384 x 1303 x 33 mm','Peso': '37,5 kg','Imagem_gerador': 'Imagens/Trina_gerador.jpg'}
gerador_hanersun_585 = {'Fabricante': 'Hanersun','SIGLA': 'HN18-72H585','Tec_construcao': 'Monocristalino','Garantia': '12 anos','Pot_max': '585 W','Eficiencia': '22,65%','Tensao_nom': '40,80 V','Tensao_aberto': '49,30 V','Corrente_nom': '10,91 A','Corrente_cc': '11,53 A','axlxp': '2278*1134*30mm','Peso': '28,5 kg','Imagem_gerador':'Imagens/Trina_gerador.jpg'}
gerador_hanersun_610 = {'Fabricante': 'Hanersun','SIGLA': 'HN21RN-66HT610W','Tec_construcao': 'Monocristalino','Garantia': '12 anos','Pot_max': '610 W','Eficiencia': '22,6 %','Tensao_nom': '40,59 V','Tensao_aberto': '48,72 V','Corrente_nom': '15,03 A','Corrente_cc': '15,94 A','axlxp': '2382 x 1134 x 30 mm','Peso': '33,5 kg','Imagem_gerador': 'Imagens/Trina_gerador.jpg'}
gerador_honor_700 = {'Fabricante': 'Honor Solar','SIGLA': 'HY-M12/132G-700','Tec_construcao': 'Monocristalino','Garantia': '12 anos','Pot_max': '700 W','Eficiencia': '22,5 %','Tensao_nom': '41,78 V','Tensao_aberto': '49,83 V','Corrente_nom': '16,77 A','Corrente_cc': '17,82 A','axlxp': '2384 x 1303 x 33 mm','Peso': '37,5 kg','Imagem_gerador': 'Imagens/Trina_gerador.jpg'}
gerador_znshine_700 = {'Fabricante': 'ZNSHINE SOLAR','SIGLA': 'ZXM8-GPLD132-700W','Tec_construcao': 'Monocristalino N-Type TOPCon Double Glass','Garantia': '12 anos','Pot_max': '700 W','Eficiencia': '22,53 %','Tensao_nom': '40,40 V','Tensao_aberto': '48,20 V','Corrente_nom': '17,33 A','Corrente_cc': '18,32 A','axlxp': '2384 x 1303 x 35 mm','Peso': '38,5 kg','Imagem_gerador': 'Imagens/Trina_gerador.jpg'}
gerador_gokin_610 = {'Fabricante': 'GOKIN SOLAR','SIGLA': 'GK-4-66HTBD-610M-F','Tec_construcao': 'Monocristalino N-Type TOPCon Dual Glass (Bifacial)','Garantia': '15 anos','Pot_max': '610 W','Eficiencia': '22,6 %','Tensao_nom': '40,22 V','Tensao_aberto': '48,00 V','Corrente_nom': '15,18 A','Corrente_cc': '16,07 A','axlxp': '2382 x 1134 x 30 mm','Peso': '33,0 kg','Imagem_gerador': 'Imagens/Trina_gerador.jpg'}
gerador_trina_400 = {'Fabricante':'Trina Solar','SIGLA':'TSM-400DE09','Tec_construcao':'Monocristalino','Garantia':'12 anos','Pot_max':'400 W','Eficiencia':'20,8 %','Tensao_nom':'34,2 V','Tensao_aberto':'41,2 V','Corrente_nom':'11,70 A','Corrente_cc':'12,28 A','axlxp':'1754 x 1096 x 30 mm','Peso':'21,0 kg','Imagem_gerador':'Imagens/Trina_gerador.jpg'}
gerador_risen_700 = {'Fabricante':'Risen Energy','SIGLA':'RSM132-8-700BHDG','Tec_construcao':'Heterojunção (HJT)','Garantia':'15 anos','Pot_max':'700 W','Eficiencia':'22,5 %','Tensao_nom':'41,78 V','Tensao_aberto':'49,83 V','Corrente_nom':'16,77 A','Corrente_cc':'17,82 A','axlxp':'2384 x 1303 x 33 mm','Peso':'37,5 kg','Imagem_gerador':'Imagens/Trina_gerador.jpg'}

inversor_growatt2 = {'Fabricante_sigla': 'Growatt NEO 2000M-X','Entradas': '4','Monitoramento': 'SIM – Wireless','Pot_nom_max': '2 kW','Tensao_nom_freq': '220 V - 54/65 Hz','Tensao_max': '65 VCC','Tensao_saida': '160 – 285 V','Corrente_max_saida': '9,3 A','Eficiencia_max': '96,5 %','axlxp_inv': '396 × 300 × 45 mm','Peso_inv': '5 kg','Nome_inversor': 'Growatt','Link_inversor': 'https://server.growatt.com/login','App_inversor': 'ShinePhone','Imagem_inversor': 'Imagens/inversor_growatt_2kw.jpg'}
inversor_growatt225 = {'Fabricante_sigla':'Growatt NEO 2250M-X2','Entradas':'4','Monitoramento':'SIM – Wireless','Pot_nom_max':'2.25 kW','Tensao_nom_freq':'35 V - 50/60 Hz','Tensao_max':'60 VCC','Tensao_saida':'220 V','Corrente_max_saida':'10,23 A','Eficiencia_max':'96,5%','axlxp_inv':'396*270*45 mm','Peso_inv':'5,1 kg','Nome_inversor':'Growatt','Link_inversor':'https://server.growatt.com/login','App_inversor':'ShinePhone','Imagem_inversor':'Imagens/inversor_growatt.jpg'}
inversor_Sungrow = {'Fabricante_sigla':'Sungrow SG3K-S','Entradas':'1','Monitoramento':'SIM – Wireless','Pot_nom_max':'3 kW','Tensao_nom_freq':'220 V - 60 Hz','Tensao_max':'600 VCC','Tensao_saida':'176 - 276 V','Corrente_max_saida':'13,7 A','Eficiencia_max':'98,2 %','axlxp_inv':'370 x 300 x 125 mm','Peso_inv':'8,5 kg','Imagem_inversor':'Imagens/inversor_hoymiles.jpg','Link_inversor': '','App_inversor': ''}
inversor_Hoymiles = {'Fabricante_sigla': 'Hoymiles MI-1500 / MI-700','Entradas': '2','Monitoramento': 'SIM – Wireless','Pot_nom_max': '1,2 kW / 0,7 kW','Tensao_nom_freq': '220 V - 45/65 Hz','Tensao_max': '60 Vcc','Tensao_saida': '180 - 275 V','Corrente_max_saida': '5,21 A / 3,36 A','Eficiencia_max': '96,50 % / 96,70%','axlxp_inv': '176 x 280 x 33 mm','Peso_inv': '3,75 kg','Imagem_inversor':'Imagens/inversor_hoymiles.jpg','Link_inversor': 'http://global.hoymiles.com','App_inversor': 'S-miles Enduser'}
inversor_hoymiles_2000 = {'Fabricante_sigla': 'Hoymiles HMS-2000-4T','Entradas': '4','Monitoramento': 'SIM – Wireless','Pot_nom_max': '2 kW','Tensao_nom_freq': '220 V - 54/65 Hz','Tensao_max': '65 VCC','Tensao_saida': '183 – 228 V','Corrente_max_saida': '9,22 A','Eficiencia_max': '99,8 %','axlxp_inv': '331 x 218 x 40.6 mm','Peso_inv': '5,56 kg','Nome_inversor': 'Hoymiles','Link_inversor': 'http://global.hoymiles.com','App_inversor': 'S-miles Enduser','Imagem_inversor': 'Imagens/inversor_hoymiles.jpg'}
inversor_hyxipower_m2000 = {'Fabricante_sigla': 'Hyxipower HYX-M2000-SW','Entradas': '4','Monitoramento': 'SIM – Wireless','Pot_nom_max': '2 kW','Tensao_nom_freq': '220 V - 50/60 Hz','Tensao_max': '65 VCC','Tensao_saida': '183 – 276 V','Corrente_max_saida': '9,09 A','Eficiencia_max': '96,70%','axlxp_inv': '310*236*35.5mm','Peso_inv': '5 kg','Nome_inversor': 'Hyxipower','Link_inversor': 'http://Hyxicloud.com','App_inversor': 'Hyxipower','Imagem_inversor': 'Imagens/inversor_hyxipower.jpg'}
inversor_solis6k = {'Fabricante_sigla': 'Solis S5-GR3P6K','Entradas': '2','Monitoramento': 'SIM – Wireless','Pot_nom_max': '6 kW','Tensao_nom_freq': '380 V - 50/60 Hz','Tensao_max': '1100 VCC','Tensao_saida': '380 V','Corrente_max_saida': '9.5 A','Eficiencia_max': '98,3%','axlxp_inv': '310 x 563 x 219 mm','Peso_inv': '17,8 kg','Nome_inversor': 'Solis','Link_inversor': '','App_inversor': '','Imagem_inversor': 'Imagens/inversor_solis_6kw.jpg'}
inversor_saj_r6_20k = {'Fabricante_sigla': 'SAJ R6-20K-T3-32-LV','Entradas': '3 MPPTs / 6 Strings','Monitoramento': 'SIM – Wi-Fi/Ethernet/4G','Pot_nom_max': '20 kW','Tensao_nom_freq': '220 V - 50/60 Hz','Tensao_max': '1100 VCC','Tensao_saida': '101.6 – 139.7 V (F-N)','Corrente_max_saida': '57,7 A','Eficiencia_max': '98,80 %','axlxp_inv': '473*659.4*240 mm','Peso_inv': '35,5 kg','Nome_inversor': 'SAJ Electric','Link_inversor': 'https://www.saj-electric.com', 'App_inversor': 'eSAJ Home / eSAJ Service','Imagem_inversor': 'Imagens/inversor_saj_r6.jpg'}
inversor_Hoymiles_1200 = {'Fabricante_sigla': 'Hoymiles MI-1200','Entradas': '4','Monitoramento': 'SIM – Wireless','Pot_nom_max': '1,2 kW','Tensao_nom_freq': '220 V - 45/65 Hz','Tensao_max': '60 Vcc','Tensao_saida': '180 - 275 V','Corrente_max_saida': '5,45 A','Eficiencia_max': '96,50 %','axlxp_inv': '176 x 280 x 33 mm','Peso_inv': '3,75 kg','Imagem_inversor':'Imagens/inversor_hoymiles_1200.jpg','Link_inversor': 'http://www.hoymiles.com','App_inversor': 'S-miles Enduser'}

# --- INTERFACE STREAMLIT ---
st.title("Gerador de Memorial Descritivo")

opcoes_mod = ["Trina Solar 695W", "Canadian Solar", "Astronergy 375W", "Gokin 585W", "Hanersun 585W", "Gokin 700W", "Astronergy 600W", "ZnShine 700W", "Hanersun 610W", "Honor 700W", "Gokin 610W","Trina Solar 400W","Risen 700W"]
mod_sel = st.selectbox("Módulo:", opcoes_mod)
tipo_mod = opcoes_mod.index(mod_sel) + 1

opcoes_inv = ["Growatt NEO 2250M-X2", "Sungrow", "Hoymiles MI-1500/700", "Hyxipower HYX-M2000-SW", "Solis 6 kW", "Hoymiles HMS-2000-4T", "Growatt NEO 2000M-X", "SAJ R6-20K-T3-32-LV","Hoymiles MI-1200"]
inv_sel = st.selectbox("Inversor:", opcoes_inv)
tipo_inv = opcoes_inv.index(inv_sel) + 1

opcoes_estrutura = ["Parafuso Prisioneiro","Laje","Solo"]
Estrutura_sel = st.selectbox("Tipo de estrutura:", opcoes_estrutura)
tipo_estrutura = opcoes_estrutura.index(Estrutura_sel)

opcoes_fornecimento = ["220 V","380 V"]
fornecimento_sel = st.selectbox("Tipo de fornecimento:", opcoes_fornecimento)
tipo_fornecimento = opcoes_fornecimento.index(fornecimento_sel)


#nome = st.text_input("Nome: ")
arquivo_pdf = st.file_uploader("Upload do PDF:", type="pdf")

if arquivo_pdf:
    with pdfplumber.open(arquivo_pdf) as pdf:
        first_page = pdf.pages[0]
        pdf_text = first_page.extract_text()

    def valor(label, texto):
        padrao = rf'{label}:\s*(.+)'
        encontrado = re.search(padrao, texto, re.IGNORECASE)
        return encontrado.group(1).strip()

    def valor_coordenada(label, texto):
        padrao = rf'{label}\s*(.+)'
        encontrado = re.search(padrao, texto, re.IGNORECASE | re.MULTILINE)
        return encontrado.group(1).strip()

    latitude = valor_coordenada('latitude',pdf_text)
    longitude = valor_coordenada('longitude',pdf_text)

    def gms_para_decimal(coordenada_gms_string):
        padrao= r'(\D)\s*(\d+)\s+(\d+)\s+([\d.,]+)'
        match = re.search(padrao, coordenada_gms_string, re.IGNORECASE)
        if not match: return "N/A"
        direcao = match.group(1).upper()
        grau = int(match.group(2))
        minuto = int(match.group(3))
        segundo = float(match.group(4).replace(',', '.'))
        valor_decimal = grau + (minuto / 60) + (segundo / 3600)
        if direcao in ['S', 'O', 'W']: valor_decimal *= -1
        return round(valor_decimal, 6)

    latitude_decimal = gms_para_decimal(latitude)
    longitude_decimal = gms_para_decimal(longitude)

    dicionario = {}
    i = 0
    for rotulo_item in rotulos:
        dicionario[rotulo_item]=valor(rotulos_pdf[i],pdf_text)
        i += 1

    Pot_ano = float(dicionario['Pot_nom'].replace(',','.')) * 128 * 12
    Pot_ano_rounded = int(Pot_ano // 1)
    Pot_ano_rounded = Pot_ano_rounded - (Pot_ano_rounded%100)
    dicionario['Pot_ano_rounded'] = str(Pot_ano_rounded)

    dicionario['Area_mod']= valor('Área Total dos Arranjos \(m²\)',pdf_text)
    Qtd_mod_var = int(dicionario['Qtd_mod'])
    Area_mod_total = round(float (dicionario['Area_mod'].replace(',','.')) * Qtd_mod_var,2)
    dicionario['Area_mod_total'] = Area_mod_total

    dicionario['Latitude'] = latitude_decimal
    dicionario['Longitude'] = longitude_decimal

    Bairro_partes = dicionario['Endereco'].rsplit(',',1)
    Bairro = Bairro_partes[1].strip()
    dicionario['Bairro'] = Bairro

    Pot_mensal = int((Pot_ano_rounded/12)//1)
    dicionario['Pot_mensal'] = Pot_mensal
    Pot_diaria = round(float(Pot_mensal / 30),2)
    dicionario['Pot_diaria'] = Pot_diaria
    Pot_nom_calc = round(Pot_diaria/(0.8*5.34),2)
    dicionario['Pot_nom_calc'] = Pot_nom_calc

    Nomes = dicionario['Nome'].split()
    Nome_capa = f'{Nomes[0]} {Nomes[-1]}'
    dicionario ['Nome_capa'] = Nome_capa
    Nome_login = f'{Nomes[0]}{Nomes[-1]}'
    dicionario['Nome_login'] = Nome_login
    Senha_login = f'{Nomes[0][0]}{Nomes[-1][0]}123456'.lower()
    dicionario['Senha_login'] = Senha_login

    #Tipo de estrutura
    dicionario['Estrutura_tipo'] = opcoes_estrutura[tipo_estrutura]

    #Tipo de fornecimento
    dicionario['Tipo_fornecimento'] = opcoes_fornecimento[tipo_fornecimento]

    if tipo_mod == 1:
        dicionario.update(gerador_trina)
        gerador_escolhido = gerador_trina
        imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
    elif tipo_mod == 2:
        dicionario.update(gerador_canadian)
        gerador_escolhido = gerador_canadian
        imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
    elif tipo_mod == 3:
        dicionario.update(gerador_astronergy) 
        gerador_escolhido = gerador_astronergy
        imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
    elif tipo_mod == 4:
        dicionario.update(gerador_gokin)
        gerador_escolhido = gerador_gokin
        imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
    elif tipo_mod == 5:
        dicionario.update(gerador_hanersun_585)
        gerador_escolhido = gerador_hanersun_585
        imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
    elif tipo_mod == 6:
        dicionario.update(gerador_gokin_700)
        gerador_escolhido = gerador_gokin_700
        imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
    elif tipo_mod == 7:
        dicionario.update(gerador_astronergy_600)
        gerador_escolhido = gerador_astronergy_600
        imagem_gerador_placeholder = gerador_astronergy_600['Imagem_gerador']   
    elif tipo_mod == 8:
        dicionario.update(gerador_znshine_700)
        gerador_escolhido = gerador_znshine_700
        imagem_gerador_placeholder = gerador_znshine_700['Imagem_gerador']  
    elif tipo_mod == 9:
        dicionario.update(gerador_hanersun_610)
        gerador_escolhido = gerador_hanersun_610
        imagem_gerador_placeholder = gerador_hanersun_610['Imagem_gerador']  
    elif tipo_mod == 10:
        dicionario.update(gerador_honor_700)
        gerador_escolhido = gerador_honor_700
        imagem_gerador_placeholder = gerador_honor_700['Imagem_gerador']
    elif tipo_mod == 11:
        dicionario.update(gerador_gokin_610)
        gerador_escolhido = gerador_gokin_610
        imagem_gerador_placeholder = gerador_gokin_610['Imagem_gerador']
    elif tipo_mod == 12:
        dicionario.update(gerador_trina_400)
        gerador_escolhido = gerador_trina_400
        imagem_gerador_placeholder = gerador_trina_400['Imagem_gerador']
    elif tipo_mod == 13:
        dicionario.update(gerador_risen_700)
        gerador_escolhido = gerador_risen_700
        imagem_gerador_placeholder = gerador_risen_700['Imagem_gerador']

    if tipo_inv == 1:
        dicionario.update(inversor_growatt225)
        inversor_escolhido = inversor_growatt225
        imagem_inversor_placeholder = inversor_growatt225['Imagem_inversor']
    elif tipo_inv == 2:
        dicionario.update(inversor_Sungrow)
        inversor_escolhido = inversor_Sungrow
        imagem_inversor_placeholder = inversor_Sungrow['Imagem_inversor']
    elif tipo_inv == 3:
        dicionario.update(inversor_Hoymiles)
        inversor_escolhido = inversor_Hoymiles
        imagem_inversor_placeholder = inversor_Hoymiles['Imagem_inversor']
    elif tipo_inv == 4: 
        dicionario.update(inversor_hyxipower_m2000)
        inversor_escolhido = inversor_hyxipower_m2000
        imagem_inversor_placeholder = inversor_hyxipower_m2000['Imagem_inversor']
    elif tipo_inv == 5: 
        dicionario.update(inversor_solis6k)
        inversor_escolhido = inversor_solis6k
        imagem_inversor_placeholder = inversor_solis6k['Imagem_inversor']
    elif tipo_inv == 6: 
        dicionario.update(inversor_hoymiles_2000)
        inversor_escolhido = inversor_hoymiles_2000
        imagem_inversor_placeholder = inversor_hoymiles_2000['Imagem_inversor'] 
    elif tipo_inv == 7: 
        dicionario.update(inversor_growatt2)
        inversor_escolhido = inversor_growatt2
        imagem_inversor_placeholder = inversor_growatt2['Imagem_inversor'] 
    elif tipo_inv == 8: 
        dicionario.update(inversor_saj_r6_20k)
        inversor_escolhido = inversor_saj_r6_20k
        imagem_inversor_placeholder = inversor_saj_r6_20k['Imagem_inversor']
    elif tipo_inv == 9: 
        dicionario.update(inversor_Hoymiles_1200)
        inversor_escolhido = inversor_Hoymiles_1200
        imagem_inversor_placeholder = inversor_Hoymiles_1200['Imagem_inversor']              

    Pot_nom_com_virgula = float(dicionario['Pot_nom'].replace(',','.'))
    Pot_max_lista = gerador_escolhido['Pot_max'].rsplit(' ',1)
    Pot_max_valor = float(Pot_max_lista[0])/1000
    dicionario['Pot_max_valor'] = Pot_max_valor
    dicionario['N_mod'] = round((Pot_nom_calc/Pot_max_valor),2)

    m = 0
    total_energia = 0
    for key in energia_mensal.keys():
        energia_mensal[key] = round(Pot_nom_com_virgula * mes[m] * irradiacao[m] * 0.8,2)
        total_energia += energia_mensal[key]
        m += 1

    dicionario.update(energia_mensal)
    dicionario['Total_energia'] = round(total_energia,2) 
    dicionario['Total_arredondado'] = 10 * math.ceil(total_energia/10)

    if st.button("Gerar Memorial Word"):
        try:
            doc = DocxTemplate('Memorial Descritivo - Template.docx')
            dicionario['imagem_gerador'] = InlineImage(doc, imagem_gerador_placeholder, width=Cm(4.0))
            dicionario['imagem_inversor'] = InlineImage(doc, imagem_inversor_placeholder, width=Cm(3.5))
            doc.render(dicionario)
            
            output_stream = io.BytesIO()
            doc.save(output_stream)
            output_stream.seek(0)
            
            st.download_button(
                label="Baixar Memorial",
                data=output_stream,
                file_name=f"Memorial Descritivo - {Nome_capa}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Erro: {e}")