import pandas as pd
import plotly.graph_objects as go

NOME_TABELA_TRANSFORMADA = 'evolucao_olin.xlsx'
NOME_COLUNA_NOME = 'Nome'

# 1. Ler a planilha específica
transformed_single_language = pd.read_excel(NOME_TABELA_TRANSFORMADA, sheet_name='Inglês')

# 2. Verificar se a coluna NOME_COLUNA_NOME existe
if NOME_COLUNA_NOME not in transformed_single_language.columns:
    raise KeyError(f"A coluna '{NOME_COLUNA_NOME}' não está presente no DataFrame")

# 3. Definir as colunas CEFR relevantes
cef_columns = ['Marco zero', 'A1.1', 'A1.2', 'A2.1', 'A2.2', 
              'B1.1', 'B1.2', 'B2.1', 'B2.2', 'B2+.1', 
              'B2+.2', 'C1.1', 'C1.2', 'C1.3']

# 4. Selecionar apenas as colunas existentes
cef_col_t = [col for col in cef_columns if col in transformed_single_language.columns]
transformed_single_language = transformed_single_language[
    [NOME_COLUNA_NOME] + cef_col_t
]

# 5. Extrair as colunas de colaboradores e datas
colaboradores = transformed_single_language[[NOME_COLUNA_NOME]]
dates = transformed_single_language[cef_col_t]

# 6. Converter os dados para formato longo (long format)
long_data = pd.melt(
    pd.concat([colaboradores, dates], axis=1),
    id_vars=[NOME_COLUNA_NOME],
    var_name='CEFR Level',
    value_name='Date'
)

# 7. Remover linhas com datas NaN
long_data = long_data.dropna(subset=['Date'])

# 8. Definir a ordem dos níveis CEFR, incluindo "Marco zero"
cef_order = ['Marco zero', 'A1.1', 'A1.2', 'A2.1', 'A2.2', 
             'B1.1', 'B1.2', 'B2.1', 'B2.2', 'B2+.1', 
             'B2+.2', 'C1.1', 'C1.2', 'C1.3', 'C2']

# 9. Converter 'CEFR Level' para categórico ordenado
long_data['CEFR Level'] = pd.Categorical(
    long_data['CEFR Level'],
    categories=cef_order,
    ordered=True
)

# 10. Ordenar os dados pelo colaborador e pela ordem dos níveis CEFR
long_data = long_data.sort_values(by=[NOME_COLUNA_NOME, 'CEFR Level'])

# 11. Criar figura
fig = go.Figure()

# 12. Preparar um dicionário para mapear as cores dos marcadores por indivíduo
marker_color_map = {}

# 13. Adicionar linhas para cada indivíduo
for name in long_data[NOME_COLUNA_NOME].unique():
    individual_data = long_data[long_data[NOME_COLUNA_NOME] == name].sort_values(by='CEFR Level')
    
    if name == 'Manoel Silva':
        line_segments = ['#391e70', '#949198']
        marker_colors = ['#391e70', '#391e70', '#949198']
    # Definir as cores das linhas com base na lógica original
    elif len(individual_data) == 2:
        line_segments = ['#adc22f']
        marker_colors = ['#391e70', '#adc22f']
    elif len(individual_data) >= 3:
        # Ajuste para mais de 2 pontos
        line_segments = ['#391e70'] * (len(individual_data) - 2) + ['#adc22f']
        marker_colors = ['#391e70'] * (len(individual_data) - 1) + ['#adc22f']
    else:
        # Apenas um ponto
        line_segments = []
        marker_colors = ['#391e70']
    
    # Adicionar segmentos de linhas conectando apenas pontos consecutivos
    for i in range(len(individual_data) - 1):
        fig.add_trace(
            go.Scatter(
                x=individual_data['CEFR Level'].iloc[i:i+2],
                y=[name, name],
                mode='lines',
                line=dict(width=2, color=line_segments[i] if i < len(line_segments) else '#391e70'),
                showlegend=False,
            )
        )
    
    # Armazenar as cores dos marcadores
    marker_color_map[name] = marker_colors

# 14. Adicionar marcadores após as linhas para garantir que fiquem por cima
for name in long_data[NOME_COLUNA_NOME].unique():
    individual_data = long_data[long_data[NOME_COLUNA_NOME] == name].sort_values(by='CEFR Level')
    colors = marker_color_map.get(name, ['#391e70'] * len(individual_data))
    
    hover_texts = individual_data['Date'].apply(lambda x: str(x).replace("\n", "\\n"))
    fig.add_trace(
        go.Scatter(
            x=individual_data['CEFR Level'],
            y=[name] * len(individual_data),
            mode='markers',
            marker=dict(size=8, color=colors),
            text=hover_texts,
            hoverinfo='text',
            showlegend=False,
        )
    )

# 15. Ordenar os membros da equipe
team_members = sorted(long_data[NOME_COLUNA_NOME].unique())

# 16. Atualizar layout da figura com a ordem correta das categorias
fig.update_layout(
    title='Mapa da Equipe e Linha do Tempo da Progressão dos Níveis CEFR',
    yaxis_title='Membros da Equipe',
    xaxis=dict(
        showgrid=True,
        fixedrange=False,  # Permitir zoom e pan
        type='category',
        categoryorder='array',
        categoryarray=cef_order,  # Ordem especificada dos níveis CEFR
        range=['Marco zero', 'C1.3'],  # Ajuste conforme necessário
        side='top'
    ),
    yaxis=dict(
        showgrid=False,
        automargin=True,
        autorange='reversed',
        fixedrange=False,  # Permitir zoom e pan
        type='category',
        categoryorder='array',
        categoryarray=team_members,  # Ordem alfabética dos membros
    ),
    template='plotly_white',
    height=min(800, 30 * len(team_members) + 150),
    width=1100,
    margin=dict(l=50, r=50, t=50, b=50),
)

# 17. Adicionar anotações para as datas nos marcadores
for _, row in long_data.iterrows():
    fig.add_annotation(
        x=row['CEFR Level'],
        y=row[NOME_COLUNA_NOME],
        text=row['Date'],
        showarrow=False,
        font=dict(size=10),
        align='center',
        xanchor='center',
        yshift=-10,  # Mover texto para baixo
        xshift=-20,  # Mover texto para a esquerda
    )

# 18. Exibir o gráfico
fig.show()