import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Impacto Quinquênio", layout="wide")

st.title("📊 Simulador de Impacto - Quinquênio")

# =============================
# BOTÃO NOVA CONSULTA
# =============================
if st.button("🔄 Nova Consulta"):
    st.session_state.clear()
    st.rerun()


# =============================
# UPLOAD PREVISÃO
# =============================
st.markdown(
    "<h3 style='margin-bottom:5px;'>📁 Upload - PREVISAO QUINQUENIO</h3>",
    unsafe_allow_html=True
)
arquivo_prev = st.file_uploader("Selecione o arquivo", type=["xlsx"], key="prev")

st.markdown("<br>", unsafe_allow_html=True)

# =============================
# UPLOAD FOLHA
# =============================
st.markdown(
    "<h3 style='margin-bottom:5px;'>📁 Upload - ARQUIVO FOLHA</h3>",
    unsafe_allow_html=True
)
arquivo_folha = st.file_uploader("Selecione o arquivo", type=["xlsx"], key="folha")

# Botão calcular
calcular = st.button("▶️ Calcular")

EVENTOS_VALIDOS = [1, 338, 1131, 1159, 1167, 1728, 1736, 1741, 1749, 1798, 1832, 1843]

MAPA_MESES = {
    'JANEIRO': 1, 'FEVEREIRO': 2, 'MARÇO': 3, 'MARCO': 3,
    'ABRIL': 4, 'MAIO': 5, 'JUNHO': 6, 'JULHO': 7,
    'AGOSTO': 8, 'SETEMBRO': 9, 'OUTUBRO': 10,
    'NOVEMBRO': 11, 'DEZEMBRO': 12
}

def get_mes_num(competencia):
    if pd.isna(competencia):
        return 0
    comp = str(competencia).strip().upper()
    return MAPA_MESES.get(comp, 0)

def meses_restantes(competencia):
    mes = get_mes_num(competencia)
    if mes > 0:
        return 12 - mes
    try:
        return 12 - pd.to_datetime(competencia).month
    except:
        try:
            return 12 - int(str(competencia)[-2:])
        except:
            return 0

# =============================
# PROCESSAMENTO
# =============================
if calcular and arquivo_prev and arquivo_folha:

    df_prev = pd.read_excel(arquivo_prev, sheet_name="PREVISAO")
    df_folha = pd.read_excel(arquivo_folha, sheet_name="FOLHA")

    df_merge = df_prev.merge(
        df_folha,
        on='Código Funcionário',
        how='inner',
        suffixes=('_prev', '_folha')
    )

    df_merge['Nome Funcionário'] = df_merge.get('Nome Funcionário_folha')
    df_merge['Cargo'] = df_merge.get('Cargo')

    df_merge = df_merge[df_merge['Código Evento'].isin(EVENTOS_VALIDOS)]

    df_merge['VALOR_CALCULADO_PERCENTUAL'] = (
        df_merge['Valor Calculado'] * df_merge['PORCENTAGEM']
    )

    df_mensal = df_merge.groupby(
        [
            'Exercício',
            'Competência',
            'Código Funcionário',
            'Nome Funcionário',
            'Cargo'
        ],
        as_index=False
    )['VALOR_CALCULADO_PERCENTUAL'].sum()

    df_mensal.rename(columns={'VALOR_CALCULADO_PERCENTUAL': 'Valor Mensal'}, inplace=True)

    # =============================
    # ORDENAÇÃO POR MÊS CORRETO
    # =============================
    df_mensal['MES_NUM'] = df_mensal['Competência'].apply(get_mes_num)

    # =============================
    # VALOR ANUAL
    # =============================
    df_mensal['Meses Restantes'] = df_mensal['Competência'].apply(meses_restantes)
    df_mensal['Valor Anual Previsto'] = (
        df_mensal['Valor Mensal'] * df_mensal['Meses Restantes']
    )

    # Ordenar
    df_mensal = df_mensal.sort_values(by=['MES_NUM', 'Nome Funcionário'])

    df_final = df_mensal[[
        'Exercício',
        'Competência',
        'Código Funcionário',
        'Nome Funcionário',
        'Cargo',
        'Valor Mensal',
        'Valor Anual Previsto'
    ]]

    # =============================
    # FORMATAR PARA EXIBIÇÃO
    # =============================
    df_exibicao = df_final.copy()
    df_exibicao['Valor Mensal'] = df_exibicao['Valor Mensal'].map("R$ {:,.2f}".format)
    df_exibicao['Valor Anual Previsto'] = df_exibicao['Valor Anual Previsto'].map("R$ {:,.2f}".format)

    st.success("✅ Cálculo realizado com sucesso!")
    st.dataframe(df_exibicao, use_container_width=True)

    # =============================
    # DOWNLOAD
    # =============================
    def gerar_excel(df):
        output = BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Resultado')

            workbook = writer.book
            worksheet = writer.sheets['Resultado']

            formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})

            worksheet.set_column('F:G', 18, formato_moeda)
            worksheet.set_column('A:E', 20)

        return output.getvalue()

    excel = gerar_excel(df_final)

    st.download_button(
        label="📥 Baixar Resultado",
        data=excel,
        file_name="impacto_quinquenio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )