import pandas as pd
import pvlib
from pvlib.pvsystem import PVSystem, Array, FixedMount
from pvlib.location import Location
from pvlib.modelchain import ModelChain
import matplotlib.pyplot as plt
import io
import streamlit as st


#@st.cache_data
def gerar_grafico (latitude, longitude, qtd_inversores, qtd_placas, pot_inv, pot_mod):
        # 1. Coordenadas do projeto (Cidade Ocidental - GO)
    #latitude, longitude = -16.0368, -47.9825

    # 2. Download automático do clima local via satélite (PVGIS)
    print("Baixando dados meteorológicos históricos...")
    clima = pvlib.iotools.get_pvgis_tmy(latitude, longitude)[0]

    # Parâmetros da Placa 
    placa = {
        'pdc0': pot_mod,           # Potência em Watts
        'gamma_pdc': -0.0034   # Coeficiente de perda térmica (~0.34%/°C p/ placas modernas)
    }

    # Parâmetros dos Inversores (Growatt NEO 2250M-X2)
    potencia_ac_total = pot_inv * qtd_inversores
    eficiencia = 0.965         # Rendimento de 96,5%

    inversor = {
        'pdc0': potencia_ac_total / eficiencia, # Teto DC que força o "Clipping" na simulação
        'eta_inv_nom': eficiencia
    }

    # 4. Estrutura Física
    suporte = FixedMount(
        surface_tilt=15,     # Inclinação do telhado (Graus)
        surface_azimuth=0    # Orientação (0 = Norte no Hemisfério Sul)
    )

    # 5. Construir o Arranjo e o Sistema
    arranjo = Array(
        mount=suporte,
        module_parameters=placa,
        temperature_model_parameters=pvlib.temperature.TEMPERATURE_MODEL_PARAMETERS['sapm']['open_rack_glass_polymer'],
        modules_per_string=qtd_placas,
        strings=1
    )

    sistema = PVSystem(
        arrays=[arranjo], 
        inverter_parameters=inversor
    )

    # 6. Rodar a Simulação
    print("Executando cálculos de geração (com clipping e perdas térmicas)...")
    local = Location(latitude, longitude, tz='America/Sao_Paulo')

    # Forçando o uso dos modelos PVWatts para interpretar o dicionário manual
    mc = ModelChain(
        sistema, local, 
        aoi_model='physical', 
        spectral_model='no_loss',
        dc_model='pvwatts',
        ac_model='pvwatts'
    )

    mc.run_model(clima)

    # 7. Formatar Resultados
    # Pegar a saída AC (energia real injetada na rede após o corte do inversor)
    geracao_mensal = mc.results.ac.resample('ME').sum() / 1000  # Converte W para kW
    geracao_mensal = geracao_mensal.round(0).astype(int)        # Arredonda

    # Montar tabela legível
    tabela = pd.DataFrame({
        'Mês': ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
        'Geração Estimada (kWh)': geracao_mensal.values
    })

    fig, ax = plt.subplots(figsize=(6, 3))
    tabela.set_index('Mês').plot(kind='bar', ax=ax, color="#2549a4")

    plt.title("Geração Estimada")
    plt.ylabel("kWh")
    plt.xlabel("")
    plt.xticks(rotation=0)
    plt.legend().set_visible(False)
    plt.tight_layout()

    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight', pad_inches=0.1)
    buffer.seek(0)
    plt.close(fig)

    return buffer, tabela #gemini alucinou ai e deu uma justificativa merda pra retornar tabela tbm, mas a imagem é o buffer

#gerar_grafico(-16.0368, -47.9825, 3, 12, 2250, 715)

'''print("\n=== RESULTADO DA SIMULAÇÃO ===")
print(tabela.to_string(index=False))
print("-" * 30)
print(f"Total Anual:  {geracao_mensal.sum()} kWh")
print(f"Média Mensal: {geracao_mensal.mean():.0f} kWh")'''

