import pandas as pd
import streamlit as st
from pptx import Presentation
import modulos.Tratamento_bases as tb
import slides.slides as sl
from io import BytesIO

#-------------------CONFIGURAÇÃO DE LAYOUT-------------------#

st.set_page_config(
    page_title="Painel de Relatórios",
    layout="wide",
    initial_sidebar_state="collapsed"
)

#-------------------CHAMANDO AS BASES-------------------#
data_base = pd.to_datetime(tb.tratamento_faturamento()['DTFATUR']).max().strftime('%d/%m/%Y')
st.write('Data da base de dados para o relatório: ', data_base)
df_representadas = tb.tabela_representadas()['REPRESENTADA'].to_list()
#-------------------CRIANDO O PAINEL--------------------#
st.title('Painel de Relatórios')
apresentacao_criada = False

with st.form(key='apresentacao'):
    st.header('PARÂMETROS PARA CRIAÇÃO DE APRESENTAÇÃO:')
    representada = st.multiselect('Selecione as representadas:', sorted(df_representadas))
    start_date = st.text_input('Mês de início do relatório (MMAAAA):')
    end_date = st.text_input('Mês de fechamento do relatório (MMAAAA):')
    lim_table = st.number_input('Limite de linhas na tabela:', min_value=5, max_value=20, value=10)
    modelo = st.selectbox('Selecione o modelo de apresentação:', ['Representadas', 'Gerência'])
    if modelo == 'Representadas':
        template_pptx = 'Suporte/Template_Representadas.pptx'
    if modelo == 'Gerência':
        template_pptx = 'Suporte/Template_Gerencia.pptx'
        representada = df_representadas
    prs = Presentation(template_pptx)
    prs_name = st.text_input('Nome da Apresentação:')
    #-------------------SLIDES-------------------#
    st.header('MODELOS DE APRESENTAÇÕES (TEMPLATES):')
    modelo_takasago = st.checkbox('Modelo Takasago', value = False)
    # Modelo Takasago: (1) purchase_usd_kg_quarter (2)sales_kg_quarter (3) oport_highlights_quarter (4) oport_abertas_quarter (5) oport_em_aberto_total 
    # (6) oport_em_aberto_ano_vigente (7) oport_convertidas_quarter (8) purchase_usd_kg (9) sales_lowlights_gp (10) sales_highlights_gp
    mensal_lbz = st.checkbox('Mensal Lubrizol', value = False)
    modelo_gerencia = st.checkbox('Modelo Gerência', value = False)

    st.divider()
    st.header('SLIDES AVULSOS:')
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown('### Slides Diretoria:')
        crescimento_queda_top20 = st.checkbox('Clientes - Crescimento e Queda - Gerencial', value=False)
        sem_vendas = st.checkbox('Clientes sem vendas (L&PC+F&N/IND & AGRO) - Gerencial', value=False)
        oport_em_aberto_gerencial_sem_group = st.checkbox('Oportunidades Em aberto (Sem GroupBy) - Gerencial', value=False)
        oport_em_aberto_gerencial_com_group = st.checkbox('Oportunidades Em aberto (Com GroupBy) - Gerencial', value=False)
        oport_criadas_gerencial_sem_group = st.checkbox('Oportunidades Criadas (Sem GroupBy) - Gerencial', value=False)
        oport_criadas_gerencial_com_group = st.checkbox('Oportunidades Criadas (Com GroupBy) - Gerencial', value=False)
        oport_faturadas_gerencial_sem_group = st.checkbox('Oportunidades Faturadas (Sem GroupBy) - Gerencial', value=False)
        oport_faturadas_gerencial_com_group = st.checkbox('Oportunidades Faturadas (Com GroupBy) - Gerencial', value=False)
        oport_sem_amostra = st.checkbox('Análise de Envio de Amostras - Gerencial', value=False)
    
    with col2:
        st.markdown('### Slides Representadas (Quarter):' + '\n' + 'Obs: para calcular slides de quarter, o mês final deve ser o último mês de um quarter (Março, Junho, Setembro ou Dezembro).')
        purchase_kg_x_usd_quarter = st.checkbox('Purchase Quarter (real x budget) - USD & KG', value=False)
        purchase_usd_kg_over_years_quarter = st.checkbox('Purchase quarter (ultimos três anos + real x budget) - USD & KG', value=False)
        sales_kg_quarter = st.checkbox('Sales Quarter (ultimos três anos + real x budget) - KG', value = False)
        oport_highlights_quarter = st.checkbox('Oportunidades Highlights Quarter', value = False)
        oport_abertas_quarter = st.checkbox('Oportunidades Abertas dentro do Quarter', value = False)
        oport_perdidas_quarter = st.checkbox('Oportunidades Perdidas Quarter', value = False)
        oport_convertidas_quarter = st.checkbox('Oportunidades Convertidas Quarter', value = False)

    with col3:
        st.markdown('### Slides Representadas (Período):')
        purchase_kg_x_usd = st.checkbox('Purchase (real x budget) - USD & KG', value=False)
        purchase_usd_kg_over_years = st.checkbox('Purchase (ultimos três anos + real x budget) - USD & KG', value=False)
        sales_kg = st.checkbox('Sales (ultimos três anos + real x budget) - KG', value = False)
        sales_lowlights = st.checkbox('Sales Lowlights (Sem Grupo de Produto)*', value = False)
        sales_lowlights_gp = st.checkbox('Sales Lowlights (Com Grupo de Produto)', value = False)
        sales_highlights = st.checkbox('Sales Highlights (Sem Grupo de Produto)', value = False)
        sales_highlights_gp = st.checkbox('Sales Highlights (Com Grupo de Produto)', value = False)
        oport_highlights = st.checkbox('Oportunidades Highlights (Gráficos de comparativa entre anos, N° e KG)', value = False)
        oport_abertas = st.checkbox('Oportunidades Abertas dentro do período', value = False)
        oport_em_aberto_ano_vigente = st.checkbox('Oportunidades Em aberto (Ano Vigente)', value = False)
        oport_em_aberto_total = st.checkbox('Oportunidades Em aberto (Total)', value = False)
        oport_perdidas = st.checkbox('Oportunidades Perdidas (Gráfico Reproved + Canceled & Tabela Reprovados)', value = False)
        oport_convertidas = st.checkbox('Oportunidades Convertidas (Análise de faturamento)', value = False)


    submit_button = st.form_submit_button(label='Criar Apresentação')

    if submit_button:
        end_date = pd.to_datetime(end_date, format='%m%Y')
        start_date = pd.to_datetime(start_date, format ='%m%Y')
        start_date_quarter = end_date - pd.DateOffset(months=2)
        sl.slide_inicio(prs_name, end_date, start_date, prs, representada)

        if modelo_takasago:
            purchase_usd_kg_over_years_quarter = True
            purchase_kg_x_usd = True
            sales_kg_quarter = True
            oport_highlights_quarter = True
            oport_abertas_quarter = True
            oport_em_aberto_total = True
            oport_em_aberto_ano_vigente = True
            oport_convertidas_quarter = True
            sales_lowlights_gp = True
            sales_highlights_gp = True
            oport_perdidas_quarter = True

        if mensal_lbz:
            representada = ['LUBRIZOL','LIPOTEC','LIPOTEC USA']
            sl.sales_rep_kg(end_date,start_date, prs, representada)
            BUs_lbz = [['PERSONAL CARE','Sem aplicação'],['I&I','HOUSEHOLD']]
            for bu in BUs_lbz:
                sl.oport_highlights(end_date, start_date, prs, representada, 1, bu)
                sl.oport_em_aberto(end_date, start_date, prs, representada, 1, 1, bu, lim_table)
                sl.oport_abertas(end_date, start_date, prs, representada, 1, bu, lim_table)
                sl.oport_perdidas_chart_table_reproved(end_date,start_date, prs, representada, lim_table, 1, bu)
                sl.oport_convertidas(end_date, start_date, prs, representada, 1, bu, lim_table)
                sl.oport_convertida_por_quarter(start_date, prs, representada, 1, bu)
                sl.oport_ativos(end_date,prs,bu,lim_table,0)
                sl.oport_ativos(end_date,prs,bu,lim_table,1)

        if modelo_gerencia:
            crescimento_queda_top20 = True
            sem_vendas = True
            oport_em_aberto_gerencial_sem_group = True
            oport_em_aberto_gerencial_com_group = True
            oport_criadas_gerencial_sem_group = True
            oport_criadas_gerencial_com_group = True
            oport_faturadas_gerencial_sem_group = True
            oport_faturadas_gerencial_com_group = True        

        if crescimento_queda_top20:
            sl.crescimento_queda_top20(end_date, prs)
        if sem_vendas:
            sl.sem_faturamento(end_date, prs, lim_table)
        if oport_em_aberto_gerencial_sem_group:
            sl.oport_em_aberto_gerencial(end_date, prs, 0, lim_table)
        if oport_em_aberto_gerencial_com_group:
            sl.oport_em_aberto_gerencial(end_date, prs, 1, lim_table)
        if oport_criadas_gerencial_sem_group:
            sl.oport_abertas_periodo(end_date, start_date, prs, 0, 1, lim_table)
        if oport_criadas_gerencial_com_group:
            sl.oport_abertas_periodo(end_date, start_date, prs, 1, 1, lim_table)
        if oport_faturadas_gerencial_sem_group:
            sl.oport_convertidas_gerencial(end_date, start_date, prs, 0, 1, lim_table)
        if oport_faturadas_gerencial_com_group:
            sl.oport_convertidas_gerencial(end_date, start_date, prs, 1, 1, lim_table)
        if oport_sem_amostra:
            sl.oport_sem_amostra_gerencial(end_date, prs)
        if purchase_kg_x_usd:
            sl.purch_rep_dolar_kg(end_date, start_date, prs, representada)
        if purchase_kg_x_usd_quarter:
            sl.purch_rep_dolar_kg(end_date, start_date_quarter, prs, representada)
        if purchase_usd_kg_over_years:
            sl.purch_rep_dolar_kg_year(end_date, start_date, prs, representada)
        if purchase_usd_kg_over_years_quarter:
            sl.purch_rep_dolar_kg_year(end_date, start_date_quarter, prs, representada)
        if sales_kg:
            sl.sales_rep_kg(end_date,start_date, prs, representada)
        if sales_kg_quarter:
            sl.sales_rep_kg(end_date,start_date_quarter, prs, representada)
        if sales_lowlights:
            sl.sales_lowlights(end_date,start_date,prs,representada,0, lim_table)
        if sales_lowlights_gp:
            sl.sales_lowlights(end_date,start_date,prs,representada,1, lim_table)
        if sales_highlights:
            sl.sales_highlights(end_date,start_date,prs,representada,0, lim_table)
        if sales_highlights_gp:
            sl.sales_highlights(end_date,start_date,prs,representada,1, lim_table)
        if oport_highlights:
            sl.oport_highlights(end_date, start_date, prs, representada,0,0)
        if oport_highlights_quarter:
            sl.oport_highlights(end_date, start_date_quarter, prs, representada,0,0)
        if oport_abertas:
            sl.oport_abertas(end_date, start_date, prs, representada,0,0, lim_table)
        if oport_abertas_quarter:
            sl.oport_abertas(end_date, start_date_quarter, prs, representada,0,0, lim_table)
        if oport_em_aberto_ano_vigente:
            sl.oport_em_aberto(end_date, start_date, prs, representada, 1,0,0, lim_table)
        if oport_em_aberto_total:
            sl.oport_em_aberto(end_date, start_date, prs, representada, 0,0,0, lim_table)
        if oport_perdidas:
            sl.oport_perdidas_chart_table_reproved(end_date,start_date, prs, representada, lim_table,0,0)
        if oport_perdidas_quarter:
            sl.oport_perdidas_chart_table_reproved(end_date,start_date_quarter, prs, representada, lim_table,0,0)
        if oport_convertidas:
            sl.oport_convertidas(end_date,start_date, prs, representada,0,0,lim_table)
        if oport_convertidas_quarter:
            sl.oport_convertidas(end_date,start_date_quarter, prs, representada,0,0,lim_table)
            
        sl.slide_fim(prs)
        binary_output = BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)
        apresentacao_criada = True
        st.success('Apresentação criada com sucesso!')

if apresentacao_criada:
    with open(prs_name+'.pptx', 'wb') as f:
        st.download_button(label='Clique aqui para baixar a apresentação', data=binary_output, file_name=prs_name+'.pptx', mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')
else:
    st.info('Crie uma apresentação para baixar o arquivo.')

st.divider()
st.header('Descrição de Modelos:')

descricao = '''**Modelo Takasago:**
- (1) purchase_usd_kg_over_years_quarter
- (2) purchase_kg_x_usd
- (3) sales_kg_quarter
- (4) oport_highlights_quarter
- (5) oport_abertas_quarter
- (6) oport_em_aberto_total
- (7) oport_em_aberto_ano_vigente
- (8) oport_convertidas_quarter
- (9) sales_lowlights_gp
- (10) sales_highlights_gp
- (11) oport_perdidas_quarter


**Mensal Lubrizol:**
- (1) sales_rep_kg
- (2) oport_highlights
- (3) oport_em_aberto
- (4) oport_abertas
- (5) oport_perdidas
- (6) oport_convertidas
- (7) oport_convertida_por_quarter

**Modelo Gerência:**
- (1) crescimento_queda_top20
- (2) sem_vendas
- (3) oport_em_aberto_gerencial_sem_group
- (4) oport_em_aberto_gerencial_com_group
- (5) oport_criadas_gerencial_sem_group
- (6) oport_criadas_gerencial_com_group
- (7) oport_faturadas_gerencial_sem_group
- (8) oport_faturadas_gerencial_com_group
'''

st.write(descricao)