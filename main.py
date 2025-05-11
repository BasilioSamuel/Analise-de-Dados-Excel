import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Carregar os dados do arquivo Excel
df = pd.read_excel('Cadastro Produtos.xlsx', sheet_name='Produtos')

# Configurar o estilo dos gráficos
sns.set_theme(style="whitegrid")
plt.figure(figsize=(12, 8))

# 1. Tabela de Participação das Marcas no Estoque
marcas_count = df['Marca'].value_counts().reset_index()
marcas_count.columns = ['Marca', 'Quantidade']
marcas_count['Participação (%)'] = (marcas_count['Quantidade'] / marcas_count['Quantidade'].sum()) * 100
marcas_count = marcas_count.sort_values(by='Quantidade', ascending=False)

# Exibir a tabela formatada
print("Tabela de Participação das Marcas no Estoque:")
print(marcas_count.to_string(index=False))

# 2. Gráfico de Barras para Participação das Marcas (Top 10)
plt.figure(figsize=(12, 6))
sns.barplot(x='Quantidade', y='Marca', data=marcas_count.head(10), palette="viridis")
plt.title('Top 10 Marcas com Maior Participação no Estoque', fontsize=16, fontweight='bold')
plt.xlabel('Quantidade de Produtos', fontsize=12)
plt.ylabel('Marca', fontsize=12)
plt.show()

# 3. Tabela de Preço Médio por Tipo de Produto
preco_medio_por_tipo = df.groupby('Tipo do Produto')['Preço Unitario'].mean().reset_index()
preco_medio_por_tipo = preco_medio_por_tipo.sort_values(by='Preço Unitario', ascending=False)

# Exibir a tabela formatada
print("\nTabela de Preço Médio por Tipo de Produto:")
print(preco_medio_por_tipo.to_string(index=False))

# 4. Gráfico de Barras para Preço Médio por Tipo de Produto
plt.figure(figsize=(12, 6))
sns.barplot(x='Preço Unitario', y='Tipo do Produto', data=preco_medio_por_tipo, palette="magma")
plt.title('Preço Médio por Tipo de Produto (R$)', fontsize=16, fontweight='bold')
plt.xlabel('Preço Médio (R$)', fontsize=12)
plt.ylabel('Tipo do Produto', fontsize=12)
plt.show()

# 5. Tabela de Margem de Lucro por Tipo de Produto
df['Margem de Lucro'] = df['Preço Unitario'] - df['Custo Unitario']  # Calcular margem de lucro
margem_por_tipo = df.groupby('Tipo do Produto')['Margem de Lucro'].mean().reset_index()
margem_por_tipo = margem_por_tipo.sort_values(by='Margem de Lucro', ascending=False)

# Exibir a tabela formatada
print("\nTabela de Margem de Lucro por Tipo de Produto:")
print(margem_por_tipo.to_string(index=False))

# 6. Gráfico de Barras para Margem de Lucro por Tipo de Produto
plt.figure(figsize=(12, 6))
sns.barplot(x='Margem de Lucro', y='Tipo do Produto', data=margem_por_tipo, palette="coolwarm")
plt.title('Margem de Lucro por Tipo de Produto (R$)', fontsize=16, fontweight='bold')
plt.xlabel('Margem de Lucro (R$)', fontsize=12)
plt.ylabel('Tipo do Produto', fontsize=12)
plt.show()