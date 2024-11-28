import pandas as pd
import plotly.graph_objects as go
import locale

# Definir o locale para português
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')  # Para sistemas baseados em Unix
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')  # Para Windows
    except:
        print("Locale português não disponível no sistema. Os meses podem aparecer em inglês.")

# Nomes dos arquivos
NOME_TABELA_TRANSFORMADA = "./saida.xlsx"

# Nome da planilha (sheet) no arquivo Excel transformado
sheet_name = 'Inglês'  # Substitua pelo nome correto da planilha
COLABORADOR = "Nome do colaborador"

# Ler os dados transformados
transformed_single_language = pd.read_excel(NOME_TABELA_TRANSFORMADA, sheet_name=sheet_name)

# Verificar se a coluna 'Nome do colaborador' existe
if COLABORADOR not in transformed_single_language.columns:
    raise KeyError(f"A coluna '{COLABORADOR}' não está presente no DataFrame")

# Extrair as colunas relevantes
cef_columns = ['MZ', 'A1', 'A1.1', 'A1.2', 'A2', 'A2.1', 'A2.2',
               'B1', 'B1.1', 'B1.2', 'B2', 'B2.1', 'B2.2', 'B2+',
               'B2+.1', 'B2+.2', 'C1', 'C1.1', 'C1.2', 'C2']
# Manter apenas as colunas existentes no DataFrame
existing_cef_columns = [col for col in cef_columns if col in transformed_single_language.columns]
transformed_single_language = transformed_single_language[[COLABORADOR] + existing_cef_columns]

# Lista completa de colaboradores
all_colaboradores = transformed_single_language[COLABORADOR].unique()

# Extrair a coluna 'Nome do colaborador' e as datas
colaboradores = transformed_single_language[[COLABORADOR]]
dates = transformed_single_language[existing_cef_columns]

# Transformar os dados para o formato longo
long_data = pd.melt(
    pd.concat([colaboradores, dates], axis=1),
    id_vars=[COLABORADOR],
    var_name='CEFR Level',
    value_name='Date'
)

# Remover linhas onde 'Date' é NaN
long_data_non_null = long_data.dropna(subset=['Date'])

# Ordenar os dados por 'Nome do colaborador' e 'CEFR Level'
long_data_non_null['CEFR Level'] = pd.Categorical(long_data_non_null['CEFR Level'], categories=cef_columns, ordered=True)
long_data_non_null = long_data_non_null.sort_values(by=[COLABORADOR, 'CEFR Level'])

# Função para definir as cores dos pontos e linhas
def mark_points(group):
    group = group.reset_index(drop=True)
    n_points = len(group)
    
    # Inicializar cores
    group['Point Color'] = '#391e70'  # Cor padrão: roxo
    group['Line Color'] = '#391e70'   # Cor padrão: roxo

    if n_points >= 1:
        # Primeiro ponto é sempre roxo
        group.at[0, 'Point Color'] = '#391e70'  # Roxo

    if n_points == 2:
        # Se tiver 2 pontos, o segundo é verde
        group.at[1, 'Point Color'] = '#adc22f'  # Verde
    elif n_points >= 3:
        # Se tiver 3 ou mais pontos, os últimos 2 são verdes
        group.at[n_points - 2, 'Point Color'] = '#adc22f'  # Verde
        group.at[n_points - 1, 'Point Color'] = '#adc22f'  # Verde

    # Definir as cores das linhas
    for idx in range(n_points - 1):
        # Se ambos os pontos são roxos, a linha é roxa
        if group.at[idx, 'Point Color'] == '#391e70' and group.at[idx + 1, 'Point Color'] == '#391e70':
            group.at[idx, 'Line Color'] = '#391e70'  # Roxo
        else:
            # Se pelo menos um dos pontos é verde, a linha é verde
            group.at[idx, 'Line Color'] = '#adc22f'  # Verde

    return group

# Aplicar a função a cada 'Nome do colaborador'
long_data_non_null = long_data_non_null.groupby(COLABORADOR, group_keys=False).apply(mark_points)

# Processar a coluna 'Date' para mostrar apenas mês e ano em português
def process_date(value):
    if isinstance(value, str):
        # Remover 'MF' ou 'Meta Final' se presente
        value = value.replace('MF', '').replace('Meta Final', '').strip()
        # Tentar converter para datetime
        try:
            date = pd.to_datetime(value, errors='coerce', dayfirst=True)
            if pd.isnull(date):
                return value  # Retorna o valor original se não for uma data válida
        except:
            return value
    elif isinstance(value, pd.Timestamp):
        date = value
    else:
        return str(value)
    # Retornar mês e ano em português
    return date.strftime('%b/%Y')  # Formato: 'Jan/2022' em português

# Aplicar a função 'process_date' à coluna 'Date'
long_data_non_null['Date_Text'] = long_data_non_null['Date'].apply(process_date)

# Criar a figura
fig = go.Figure()

# Adicionar traços para cada 'Nome do colaborador'
for name in all_colaboradores:
    group = long_data_non_null[long_data_non_null[COLABORADOR] == name]
    if group.empty:
        # Se o colaborador não tem dados, adicionar um traço vazio
        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[name],
                mode='markers',
                marker=dict(size=8, color='rgba(0,0,0,0)'),  # Marcador transparente
                showlegend=False,
                hoverinfo='none'
            )
        )
    else:
        group = group.sort_values('CEFR Level')
        x_values = group['CEFR Level']
        y_values = [name] * len(group)
        text_values = group['Date_Text']
        marker_colors = group['Point Color'].tolist()
        line_colors = group['Line Color'].tolist()

        # Adicionar pontos e linhas
        for i in range(len(group) - 1):
            fig.add_trace(
                go.Scatter(
                    x=x_values.iloc[i:i+2],
                    y=y_values[i:i+2],
                    mode='lines+markers+text',
                    marker=dict(size=8, color=marker_colors[i:i+2]),
                    line=dict(width=2, color=line_colors[i]),
                    text=text_values.iloc[i:i+2],
                    textposition="top center",
                    hoverinfo='text',
                    showlegend=False,
                )
            )
        # Adicionar o último ponto
        fig.add_trace(
            go.Scatter(
                x=[x_values.iloc[-1]],
                y=[y_values[-1]],
                mode='markers+text',
                marker=dict(size=8, color=marker_colors[-1]),
                text=[text_values.iloc[-1]],
                textposition="top center",
                hoverinfo='text',
                showlegend=False,
            )
        )

# Definir a ordem das categorias no eixo x
fig.update_xaxes(
    type='category',
    categoryorder='array',
    categoryarray=cef_columns,
    title_text='Nível CEFR',
    showgrid=True,
    side='top'
)

# Definir o eixo y com todos os colaboradores, mesmo sem dados
fig.update_yaxes(
    type='category',
    categoryorder='array',
    categoryarray=sorted(all_colaboradores),
    title_text='Colaboradores',
    showgrid=False,
    automargin=True,
    autorange='reversed',
    tickfont=dict(size=12),
)

# Ajustar a altura da figura
height_per_collaborator = 20  # Ajuste este valor para aumentar ou diminuir o espaçamento
fig_height = 600

# Atualizar o layout da figura
fig.update_layout(
    title='Progresso dos Colaboradores nos Níveis CEFR',
    template='plotly_white',
    height=fig_height,
    width=1200,
    margin=dict(l=50, r=50, t=50, b=50),
)

# Mostrar a figura
fig.show()
