import pandas as pd
import re

def process_table(input_file, output_file):
    # Ler o arquivo de entrada
    df = pd.read_excel(input_file)
    
    # Listas para armazenar colunas de datas e avaliações
    date_columns = []
    assessment_columns = []
    
    # Usar expressões regulares para identificar colunas de data e avaliação
    for col in df.columns:
        if re.match(r'Data(\.\d+)?', col):
            date_columns.append(col)
        elif re.match(r'Avaliação.*', col):
            assessment_columns.append(col)
    
    # Ordenar as colunas para garantir o mapeamento correto
    date_columns.sort()
    assessment_columns.sort()
    
    # Parear colunas de data e avaliação
    date_assessment_pairs = list(zip(date_columns, assessment_columns))
    
    # Lista ordenada dos níveis CEFR incluindo 'MZ'
    cefr_order = ['MZ', 'A1', 'A1.1', 'A1.2', 'A2', 'A2.1', 'A2.2', 'B1', 'B1.1', 'B1.2', 'B2', 'B2.1', 'B2.2', 'B2+', 'B2+.1', 'B2+.2', 'C1', 'C1.1', 'C1.2', 'C2', 'C2.1', 'C2.2']
    
    # Criar um escritor Excel para múltiplas planilhas
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # Processar cada idioma separadamente
    for language in ['Inglês', 'Espanhol']:
        lang_df = df[df['Idioma'] == language]
        
        # Conjunto para armazenar todos os níveis CEFR encontrados
        cefr_levels = set()
        
        # Dicionário para armazenar os dados dos colaboradores
        collaborator_data = {}
        
        for index, row in lang_df.iterrows():
            name = row['Nome do colaborador']
            assessments = []
    
            # Coletar pares de data e avaliação
            for date_col, assessment_col in date_assessment_pairs:
                if pd.notnull(row[date_col]) and pd.notnull(row[assessment_col]):
                    date = row[date_col]
                    level = row[assessment_col]
                    assessments.append((date, level))
                    cefr_levels.add(level)
    
            # Obter a meta final
            final_goal = row['Meta final']
            cefr_levels.add(final_goal)
    
            # Armazenar os dados do colaborador
            collaborator_data[name] = {
                'assessments': assessments,
                'final_goal': final_goal
            }
    
        # Garantir que 'MZ' esteja incluído nos níveis CEFR
        cefr_levels.add('MZ')
    
        # Ordenar os níveis CEFR encontrados de acordo com a ordem especificada
        cefr_levels = [level for level in cefr_order if level in cefr_levels]
    
        # Construir o DataFrame de saída
        output_columns = ['Nome do colaborador'] + cefr_levels
        output_df = pd.DataFrame(columns=output_columns)
    
        # Preencher os dados no DataFrame de saída
        for name, data in collaborator_data.items():
            row_dict = {'Nome do colaborador': name}
            assessments = data['assessments']
            final_goal = data['final_goal']
    
            # Inserir datas das avaliações nas células correspondentes
            for date, level in assessments:
                date_str = date.strftime('%Y-%m-%d') if isinstance(date, pd.Timestamp) else str(date)
                if level in row_dict and pd.notnull(row_dict[level]):
                    pass
                else:
                    row_dict[level] = date_str.lower()
    
            # Inserir a meta final na célula correspondente
            if final_goal in row_dict and pd.notnull(row_dict[final_goal]):
                pass
            else:
                row_dict[final_goal] = 'MF'
    
            # Adicionar a linha ao DataFrame de saída
            output_df = output_df._append(row_dict, ignore_index=True)
    
        # Reordenar as colunas
        output_df = output_df[output_columns]
    
        # Escrever o DataFrame na planilha correspondente
        output_df.to_excel(writer, sheet_name=language, index=False)
    
    # Salvar o arquivo Excel com múltiplas planilhas
    writer._save()

# Exemplo de uso
process_table('Dados para mapeamento.xlsx', 'saida.xlsx')
