import pandas as pd

# Caminho do arquivo Excel original
file_path = r"C:\Users\FDR Thay\Downloads\Diagnóstico_do_Projeto_II_Etapa2024-06-17_07_38_32.xlsx"

# Carregar os dados do arquivo Excel
df = pd.read_excel(file_path)

# Colunas relevantes para a nova tabela
colunas_relevantes = ['Submission Date',
                      'Professor',
                      'E-mail',
                      'Patrocinador',
                      'Cidade',
                      'NOME COMPLETO DO ALUNO',
                      'ESCOLA/NÚCLEO:',
                      'TURNO:',
                      'SÉRIE/ANO:',
                      'IDADE:',
                      'Local de execução',
                      '03 - Você gosta de praticar esportes?',
                      '03.1 Quais esportes você mais gosta?',
                      '04. Você considera que tem alguma dificuldade para praticar esportes?',
                      '05 - Você sabe o que é Fair Play?',
                      '06 - Marque as opções de Fair Play que você já pratica:',
                      '07 - Você sabe o que é protagonismo?',
                      '08 - O que você considera ser protagonismo?',
                      '10 - Quais valores você tem praticado até o momento?',
                      '11 - Você gosta de futebol?',
                      '14 - Você conhece os dribles do futebol?',
                      '15 - Quando você enfrenta dificuldades no dia a dia, como acha que pode superá-las? Marque as opções que você consegue fazer:',
                      '16 - Como o Projeto Futebol de Rua Pela Educação pode ajudar no seu desenvolvimento? Marque as opções mais importantes para você:',
                      '17 - Quais dos temas a seguir você já sabe ou já estudou antes de participar do projeto?']

# Selecionar apenas as colunas relevantes
df_relevante = df[colunas_relevantes]

# Identifica duplicatas baseando-se em colunas específicas (Nome e Idade)
# Mantém apenas uma das duplicatas
df_relevante = df_relevante.drop_duplicates(subset=['NOME COMPLETO DO ALUNO', 'IDADE:'], keep='first')

# Função para contabilizar as respostas e marcar as três mais comuns
def contar_respostas(coluna, n=3):
    # Separar as respostas em itens individuais e contar a frequência de cada um
    todas_respostas = df_relevante[coluna].dropna().str.split(', ').explode()
    frequencias = todas_respostas.value_counts()
    
    # Identificar as três respostas mais comuns
    top_n_respostas = frequencias.nlargest(n).index.tolist()
    
    # Marcar as outras respostas como "Outros"
    def marcar_respostas(respostas):
        marcadas = [resposta if resposta in top_n_respostas else 'Outros' for resposta in respostas.split(', ')]
        return ', '.join(marcadas)
    
    # Aplicar a marcação de "Outros" para todas as respostas na coluna
    return df_relevante[coluna].dropna().apply(marcar_respostas)

# Aplicar a função de contabilização para cada coluna específica
colunas_especificas = [
    '03.1 Quais esportes você mais gosta?',
    '06 - Marque as opções de Fair Play que você já pratica:',
    '08 - O que você considera ser protagonismo?',
    '10 - Quais valores você tem praticado até o momento?',
    '14 - Você conhece os dribles do futebol?',
    '15 - Quando você enfrenta dificuldades no dia a dia, como acha que pode superá-las? Marque as opções que você consegue fazer:',
    '16 - Como o Projeto Futebol de Rua Pela Educação pode ajudar no seu desenvolvimento? Marque as opções mais importantes para você:',
    '17 - Quais dos temas a seguir você já sabe ou já estudou antes de participar do projeto?'
]

for coluna in colunas_especificas:
    df_relevante[coluna] = contar_respostas(coluna, n=3)


# Caminho para salvar a nova planilha
output_file = r"C:\Users\FDR Thay\Downloads\tabela_atualizada.xlsx"

# Salvar a nova tabela em um arquivo Excel
df_relevante.to_excel(output_file, index=False)

print(f"Nova planilha criada e salva em {output_file}")
