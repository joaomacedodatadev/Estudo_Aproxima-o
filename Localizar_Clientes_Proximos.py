#Query que localiza clientes próximos através de coordenadas geográficas

import geopandas as gpd
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
import os

#Abrir a janela do tkinter
root = tk.Tk()
root.withdraw() #Ocultar a janela principal do tkinter

#Abrir a janela para o usuário selecionar o arquivo de entrada
print("\n Selecione o seu arquivo de superintendência...")
ca_base_geral_grupo = filedialog.askopenfilename(title="Selecione o seu arquivo de superintendência", filetypes=[("Arquivos Excel","*.xlsx")])

#Se o usuário não selecionar uma pasta, script encerra
if not ca_base_geral_grupo:
    print("\n Nenhum arquivo selecionado. O Script foi encerrado.")
    exit()

#Abrir a janela para o usuário selecionar o arquivo de entrada
print("\n Selecione o seu arquivo que deseja saber o grupo ideal...")
ca_base_inicial = filedialog.askopenfilename(title="Selecione o seu arquivo que deseja saber o grupo ideal", filetypes=[("Arquivos Excel","*.xlsx")])

#Se o usuário não selecionar uma pasta, script encerra
if not ca_base_inicial:
    print("\n Nenhum arquivo selecionado. O Script foi encerrado.")
    exit()

#Abre a jenalea para o usuário selecionar a pastsa de saída
print("\n Selecione o local de destino..")
diretorio_saida = filedialog.askdirectory(title="Selecione o local de destino")

#Se o usário não selecionar uma pasta de destino, o script encerra
if not diretorio_saida:
    print("\n Nenhuma pasta selecionada. O script foi encerrado.")
    exit()

#Caminho Fixo
ca_base_fora_rota = fr"C:\Users\{os.getlogin()}\AEGEA Saneamento e Participações S.A\Análise Qualidade - General\2.Controle_de_demandas\2. Demandas\3. Incremento_Coordenadas\0.Fora_Rota\dFR.xlsx"

# Leitura das bases
base_modelo = pd.read_excel(ca_base_geral_grupo)
base_origem = pd.read_excel(ca_base_inicial)
base_fora_rota = pd.read_excel(ca_base_fora_rota)
base_fora_rota_max = base_fora_rota[base_fora_rota['REFERENCIA'] == base_fora_rota['REFERENCIA'].max()]

# Conversão de coordenadas para pontos em mapa
geo_modelo = gpd.GeoDataFrame(base_modelo, 
                              geometry=gpd.points_from_xy(base_modelo['LONGITUDE_ORIGEM'], 
                                                          base_modelo['LATITUDE_ORIGEM']), 
                              crs=4674)

geo_origem = gpd.GeoDataFrame(base_origem,
                              geometry=gpd.points_from_xy(base_origem['LONGITUDE_DESTINO'],
                                                          base_origem['LATITUDE_DESTINO']), 
                              crs=4674)

# Conversão de coordenadas para UTM
geo_modelo = geo_modelo.to_crs(31983)
geo_origem = geo_origem.to_crs(31983)

# Encontrar o ponto mais próximo
def encontrar_ponto_mais_proximo(row, origem):
    if row.geometry is None:
        return None  # Retorna None para evitar erro
    distancias = origem.distance(row.geometry)  # Calcula as distâncias
    distancias = distancias[distancias.notna()]  # Remove distâncias NaN
    if distancias.empty:  # Se não houver nenhuma distância válida
        return None
    return distancias.idxmin()  # Retorna índice do ponto mais próximo

# Aplicar a função para encontrar o índice do ponto mais próximo
idx_mais_proximo = geo_origem.apply(lambda row: encontrar_ponto_mais_proximo(row, geo_modelo), axis=1)
idx_mais_proximo = idx_mais_proximo.dropna().astype(int)

# Criar coluna com a matrícula do ponto mais próximo
geo_origem["MATRICULA_MAIS_PROXIMA"] = geo_origem.index.map(
    lambda i: geo_modelo.loc[idx_mais_proximo[i], "NUM_LIGACAO"] if i in idx_mais_proximo else None
)

# Verificar quantos valores são NaN
print(f"Total de valores NaN: {geo_origem['MATRICULA_MAIS_PROXIMA'].isna().sum()}")

#Trazendo informações
join = pd.merge(
    geo_origem[['NUM_LIGACAO','MATRICULA_MAIS_PROXIMA','END_LIGACAO','TIPO_LIGACAO','LATITUDE_DESTINO','LONGITUDE_DESTINO']],
    geo_modelo[['NUM_LIGACAO','GRUPO','ROTA','STATUS_LIGACAO']],
    left_on='MATRICULA_MAIS_PROXIMA',
    right_on='NUM_LIGACAO',
    how='left'
)
join_FR = pd.merge(
    join,
    base_fora_rota_max[['NUM_LIGACAO','DSC_OCORRENCIA']],
    left_on='NUM_LIGACAO_x',
    right_on='NUM_LIGACAO',
    how='left'
)
# Converter o DataFrame 'join' para um GeoDataFrame
geo_join = gpd.GeoDataFrame(join_FR, 
                            geometry=gpd.points_from_xy(join['LONGITUDE_DESTINO'], 
                                                        join['LATITUDE_DESTINO']), 
                            crs=4674)
#Excluir colunas
geo_join = geo_join.drop(columns=['NUM_LIGACAO_y','NUM_LIGACAO'], axis=1)
#Renomear coluna
geo_join = geo_join.rename(columns={'NUM_LIGACAO_x': 'NUM_LIGACAO'})
#Tratamento de coordenadas
geo_join['LATITUDE_DESTINO'] = geo_join['LATITUDE_DESTINO'].astype(str).str.replace(',','.')
geo_join['LONGITUDE_DESTINO'] = geo_join['LONGITUDE_DESTINO'].astype(str).str.replace(',','.')
# Salvar o resultado corrigido
nome_saida_excel = f'Resultado_final.xlsx'
nome_saida_geo = f'Resultado_final.gpkg'
caminho_saida_excel = os.path.join(diretorio_saida,nome_saida_excel)
caminho_saida_geo = os.path.join(diretorio_saida, nome_saida_geo)
geo_join.to_file(caminho_saida_geo, driver="GPKG")
geo_join.to_excel(caminho_saida_excel,index=False)

# Plotar as duas bases no mapa
fig, ax = plt.subplots(figsize=(10, 10))
geo_modelo.plot(ax=ax, color='blue', label='Modelo', alpha=0.5)
geo_origem.plot(ax=ax, color='red', label='Origem', alpha=0.5)

# Adicionar título e legenda
plt.title('Origem e Modelo no Mapa')
plt.legend()

# Exibir o gráfico
plt.show()
