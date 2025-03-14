import pandas as pd
from pathlib import Path
from scipy.optimize import minimize

# Caminho
input_dir = Path("./input").resolve()
output_dir = Path("./output").resolve()

# Lendo o arquivo
df = pd.read_excel(input_dir / "Auditoria_2025-02_28_Total.xlsx")

# Mantém apenas essas colunas
df = df[["DataFornecimento", "EmpresaResponsavel", "Resultado", "resultadpVpl"]] 

# Conferência de valores:
soma_resultado = df['Resultado'].sum()
soma_resultadovpl = df['resultadpVpl'].sum()
print(soma_resultado)
print(soma_resultadovpl)

# Agrupar meses para cálculo:
agrupamento = df.groupby(['DataFornecimento', 'EmpresaResponsavel'], as_index=False).sum()
print(agrupamento)
df_agrupado = agrupamento
df2 = df_agrupado.copy()

# Calculando taxa media por empresa:
def processar_dataframe(df2):
    # Função objetivo que minimiza a taxa inicial
    def objetivo(tx_inicial):
        return tx_inicial[0]

    # Função de restrição para que a diferença entre as somas usando valores originais seja zero
    def restricao(tx_inicial):
        df2['tx_inicial'] = tx_inicial[0]
        df2['mult_inicial'] = (1 + df2['tx_inicial']).cumprod()
        
        # Calcular as colunas `vpl_taxa_inicial` com valores originais
        df2['vpl_taxa_inicial'] = df2['Resultado'] / df2['mult_inicial']
        
        # Restrições de igualdade: queremos que essa diferença seja zero
        return df2['resultadpVpl'].sum() - df2['vpl_taxa_inicial'].sum()

    # Definindo a restrição como uma igualdade
    restricao_eq = {'type': 'eq', 'fun': restricao}

    # Executando o solver com um método específico e uma tolerância mais alta
    resultado_solver = minimize(
        objetivo, 
        x0=[0.005], 
        bounds=[(0, 1)],  # Permitindo taxas negativas e positivas
        constraints=[restricao_eq],
        method='SLSQP',  # Algoritmo que suporta restrições e pode ser mais robusto para este problema
        options={'ftol': 1e-9, 'maxiter': 1000}  # Maior tolerância e número máximo de iterações
    )
    taxa_otima = resultado_solver.x[0]

    # Aplicando a taxa ótima encontrada no DataFrame
    df2['tx_inicial'] = taxa_otima
    df2['mult_inicial'] = (1 + df2['tx_inicial']).cumprod()
    
    # Calculando as colunas finais de `vpl_taxa_inicial` e `vpl_taxa_inicial_abs`
    df2['vpl_taxa_inicial'] = df2['Resultado'] / df2['mult_inicial']
    df2['vpl_taxa_inicial_abs'] = df2['Resultado'].abs() / df2['mult_inicial']
    
    return df2

# Dividindo o DataFrame original em df_x e df_y
df_x = df2[df2['EmpresaResponsavel'] == 'SAFIRA ADMINISTRACAO E COMERCIALIZACAO DE ENERGIA S.A.'].copy()
df_y = df2[df2['EmpresaResponsavel'] == 'SAFIRA VAREJO COMERCIALIZACAO DE ENERGIA LTDA'].copy()
df_z = df2[df2['EmpresaResponsavel'] == 'SAFIRA ARTEMIS COMERCIALIZACAO DE ENERGIA LTDA'].copy()

# Aplicando a função a ambos os DataFrames
df_x = processar_dataframe(df_x)
df_y = processar_dataframe(df_y)
df_z = processar_dataframe(df_z)

# Resultado final de cada DataFrame
print("DataFrame para Empresa ADM:")
print(df_x)
print("\nDataFrame para Empresa VRJ:")
print(df_y)
print("\nDataFrame para Empresa VRJ:")
print(df_z)

# Juntar as empresas em um único dataframe:
def concat_dataframes(df_x, df_y, df_z, axis=0):
    """
    Concatena dois DataFrames ao longo do eixo especificado e retorna um novo DataFrame.
    
    Parâmetros:
        df1 (pd.DataFrame): O primeiro DataFrame.
        df2 (pd.DataFrame): O segundo DataFrame.
        axis (int): O eixo ao longo do qual concatenar. 0 para linhas, 1 para colunas. Padrão é 0.
        
    Retorna:
        pd.DataFrame: O novo DataFrame concatenado.
    """
    novo_df = pd.concat([df_x, df_y,df_z], axis=axis, ignore_index=True if axis == 0 else False)
    return novo_df

novo_df = concat_dataframes(df_x, df_y, df_z, axis=0)
novo_df

# Função para processar o DataFrame consolidado
def processar_dataframe(df2):
    # Função objetivo que minimiza a taxa inicial
    def objetivo(tx_inicial):
        return tx_inicial[0]

    # Função de restrição para que a diferença entre as somas usando valores originais seja zero
    def restricao(tx_inicial):
        df2['tx_inicial'] = tx_inicial[0]
        df2['mult_inicial'] = (1 + df2['tx_inicial']).cumprod()
        
        # Calcular as colunas `vpl_taxa_inicial` com valores originais
        df2['vpl_taxa_inicial'] = df2['Resultado'] / df2['mult_inicial']
        
        # Restrições de igualdade: queremos que essa diferença seja zero
        return df2['resultadpVpl'].sum() - df2['vpl_taxa_inicial'].sum()

    # Definindo a restrição como uma igualdade
    restricao_eq = {'type': 'eq', 'fun': restricao}

    # Executando o solver com um método específico e uma tolerância mais alta
    resultado_solver = minimize(
        objetivo, 
        x0=[0.005], 
        bounds=[(0, 1)],  # Permitindo taxas entre 0 e 1
        constraints=[restricao_eq],
        method='SLSQP',  # Algoritmo que suporta restrições e pode ser mais robusto para este problema
        options={'ftol': 1e-9, 'maxiter': 1000}  # Maior tolerância e número máximo de iterações
    )
    taxa_otima = resultado_solver.x[0]

    # Aplicando a taxa ótima encontrada no DataFrame
    df2['tx_inicial'] = taxa_otima
    df2['mult_inicial'] = (1 + df2['tx_inicial']).cumprod()
    
    # Calculando as colunas finais de `vpl_taxa_inicial` e `vpl_taxa_inicial_abs`
    df2['vpl_taxa_inicial'] = df2['Resultado'] / df2['mult_inicial']
    df2['vpl_taxa_inicial_abs'] = df2['Resultado'].abs() / df2['mult_inicial']
    
    return df2

# Consolidando os dados por Data (somando as colunas relevantes)
df_consolidado = df.groupby('DataFornecimento').agg({
    'Resultado': 'sum',
    'resultadpVpl': 'sum'
}).reset_index()

# Aplicando a função ao DataFrame consolidado
df_consolidado = processar_dataframe(df_consolidado)

# Resultado final
print("DataFrame consolidado com taxa única:")
print(df_consolidado)

# Adicionando a coluna "Empresa" com o valor fixo
df_consolidado['EmpresaResponsavel'] = 'ADM + VAREJO + ARTEMIS'

# Resultado final
print("DataFrame consolidado com taxa única e coluna Empresa:")
print(df_consolidado)

df_consolidado.to_excel(output_dir / "taxa_media_consolidado.xlsx", index=False)
novo_df.to_excel(output_dir / "taxa_media.xlsx", index=False)