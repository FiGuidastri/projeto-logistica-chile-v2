import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import io # Usado para manipular o arquivo em memória
import os # Usado para manipular nomes de arquivos

# =====================================================================================
#  NOSSA LÓGICA DE REPROGRAMAÇÃO (ADAPTADA DENTRO DE UMA FUNÇÃO)
# =====================================================================================
def processar_planilha(arquivo_excel, dia_feriado):
    """
    Esta função contém o nosso algoritmo principal.
    Ela recebe o arquivo carregado e o dia do feriado, e retorna
    o arquivo modificado e uma lista de logs.
    """
    logs = [] # Lista para armazenar as mensagens de log

    try:
        # Usamos o arquivo em memória diretamente
        workbook = openpyxl.load_workbook(arquivo_excel)
        sheet_name = '01. Calendario SCL Abarrotes'
        sheet = workbook[sheet_name]
        logs.append(f"Planilha '{arquivo_excel.name}' carregada com sucesso.")
    except Exception as e:
        logs.append(f"ERRO: Não foi possível ler a planilha. Verifique se é o arquivo correto. Detalhe: {e}")
        return None, logs

    delivery_columns = ['AI', 'AJ', 'AK', 'AL', 'AM', 'AN']
    observations_column = 'CT'
    weekday_map = {'L': 1, 'M': 2, 'W': 3, 'J': 4, 'V': 5, 'S': 6, 'D': 7}

    holiday_col_letter = None
    for col_letter in delivery_columns:
        day_in_sheet = sheet[f'{col_letter}3'].value
        if day_in_sheet == dia_feriado:
            holiday_col_letter = col_letter
            break
    
    if not holiday_col_letter:
        logs.append(f"ERRO: O dia {dia_feriado} não foi encontrado na linha 3 das colunas de Entrega.")
        return None, logs
    
    logs.append(f"Feriado identificado na coluna de Entrega: {holiday_col_letter}")

    holiday_col_index = column_index_from_string(holiday_col_letter)
    if holiday_col_index == column_index_from_string(delivery_columns[0]):
         logs.append("Atenção: O feriado é o primeiro dia do período. Não é possível antecipar.")
         return None, logs

    previous_col_index = holiday_col_index - 1
    previous_col_letter = get_column_letter(previous_col_index)
    
    tasks_moved = 0
    
    for row_index in range(8, sheet.max_row + 1):
        task_cell = sheet[f'{holiday_col_letter}{row_index}']
        if isinstance(task_cell.value, (int, float)) and 1 <= task_cell.value <= 6:
            destination_cell = sheet[f'{previous_col_letter}{row_index}']
            weekday_initial = sheet[f'{previous_col_letter}6'].value.upper()
            new_weekday_number = weekday_map.get(weekday_initial)
            
            if new_weekday_number:
                destination_cell.value = new_weekday_number
                task_cell.value = None
                log_message = f"Entrega reprogramada (com substituição) do dia {dia_feriado} para a coluna {previous_col_letter}."
                sheet[f'{observations_column}{row_index}'].value = log_message
                tasks_moved += 1
    
    logs.append(f"Reprogramação concluída. {tasks_moved} tarefas foram movidas.")
    
    return workbook, logs

# =====================================================================================
#  INTERFACE WEB COM STREAMLIT
# =====================================================================================

st.title("🤖 Reprogramador Automático de Feriados")

st.write("""
Esta ferramenta automatiza a reprogramação de entregas em planilhas de logística. 
Basta carregar sua planilha, informar o dia do feriado e clicar em 'Reprogramar'.
""")

uploaded_file = st.file_uploader(
    "1. Escolha a sua planilha de programação (.xlsx)",
    type=['xlsx']
)

holiday_day = st.number_input(
    "2. Digite o dia do mês que é feriado (ex: 20)",
    min_value=1, 
    max_value=31, 
    step=1,
    value=20 # Valor padrão para facilitar o teste
)

if st.button("Reprogramar Planilha"):
    if uploaded_file is not None:
        with st.spinner('Aguarde... Reprogramando tarefas...'):
            modified_workbook, logs = processar_planilha(uploaded_file, int(holiday_day))

        st.subheader("Relatório da Operação:")
        for log in logs:
            st.info(log)
        
        if modified_workbook:
            output = io.BytesIO()
            modified_workbook.save(output)
            output.seek(0)
            
            st.success("Sua planilha foi reprogramada com sucesso!")
            
            # --- LÓGICA DO NOME DO ARQUIVO ALTERADA AQUI ---
            # 1. Pega o nome do arquivo original (ex: 'planilha.xlsx')
            original_filename = uploaded_file.name
            # 2. Separa o nome base da extensão (ex: 'planilha', '.xlsx')
            base_name, extension = os.path.splitext(original_filename)
            # 3. Cria o novo nome com o sufixo
            new_filename = f"{base_name}_reprogramado{extension}"
            
            st.download_button(
                label="Clique aqui para baixar a planilha reprogramada",
                data=output,
                # 4. Usa o novo nome do arquivo no botão de download
                file_name=new_filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    else:
        st.error("Por favor, carregue uma planilha antes de reprogramar.")