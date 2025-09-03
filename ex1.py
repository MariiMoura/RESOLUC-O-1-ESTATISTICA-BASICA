import pandas as pd #abra o terminal e baixe a biblioteca pandas com o comando: pip install pandas
import os #abra o terminal e baixe a biblioteca openpyxl com o comando: pip install openpyxl

# Tabelas originais
a = [
    [6,1,1,10,6,4,8,1,6,2],
    [1,4,10,0,4,9,5,6,4,9],
    [10,1,6,7,6,1,4,3,6,0],
    [2,4,5,5,5,2,9,6,8,3],
    [1,8,4,6,1,1,8,7,4,3]
]
b = [
    [79,93,45,77,31,74,99,54,88,11],
    [61,36,24,3,32,50,50,20,29,46],
    [6,15,54,3,49,12,63,65,66,29],
    [52,47,65,77,63,61,65,86,71,67],
    [51,13,69,30,18,59,98,34,85,93]
]
c = [
    [0.09,0.09,0.05,0.03,0.00,0.02,0.09,0.09,0.03,0.02],
    [0.04,0.07,0.01,0.09,0.09,0.02,0.07,0.08,0.09,0.04],
    [0.03,0.06,0.06,0.05,0.09,0.06,0.01,0.06,0.03,0.02],
    [0.06,0.09,0.08,0.06,0.01,0.06,0.06,0.07,0.09,0.07],
    [0.09,0.06,0.07,0.01,0.04,0.04,0.08,0.01,0.09,0.10]
]

# Criando DataFrames
df_a = pd.DataFrame(a)
df_b = pd.DataFrame(b)
df_c = pd.DataFrame(c)

# Linhas com fórmulas para Excel
extra_rows = {
    "Média": "=MÉDIA(A1:J5)",
    "Mediana": "=MED(A1:J5)",
    "Moda": "=MODA(A1:J5)",
    "Variância": "=VAR.P(A1:J5)",
    "Desvio-padrão": "=DESVPAD.P(A1:J5)"
}
df_extra = pd.DataFrame([[k, v] for k, v in extra_rows.items()], columns=["Medida", "Fórmula"])

# Resultados já calculados
dados = {
    "Exercício": ["a", "b", "c"],
    "Média": [4.66, 51.36, 0.0564],
    "Mediana": [4.5, 53.0, 0.06],
    "Moda": [1, 65, 0.09],
    "Variância": [8.3922, 711.2963, 0.000844],
    "Desvio-padrão": [2.8969, 26.6701, 0.0291]
}
df_res = pd.DataFrame(dados)

# Caminho da pasta onde está o script
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(current_dir, "estatistica_final.xlsx")

# Criando arquivo Excel com múltiplas abas
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    df_a.to_excel(writer, sheet_name="Tabela A", index=False, header=False)
    df_extra.to_excel(writer, sheet_name="Tabela A", startrow=7, index=False)

    df_b.to_excel(writer, sheet_name="Tabela B", index=False, header=False)
    df_extra.to_excel(writer, sheet_name="Tabela B", startrow=7, index=False)

    df_c.to_excel(writer, sheet_name="Tabela C", index=False, header=False)
    df_extra.to_excel(writer, sheet_name="Tabela C", startrow=7, index=False)

    df_res.to_excel(writer, sheet_name="Resultados", index=False)

print(f"Arquivo '{excel_path}' criado com sucesso!")
