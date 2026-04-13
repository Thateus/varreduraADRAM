import duckdb

con = duckdb.connect('boletim.db')
con.execute("INSTALL spatial; LOAD spatial;")

caminho = r"D:\pcp.faxinal\Documentos\py\BOLETIM DIARIO 31.xlsx"

# Puxa o bloco das linhas 12 a 17 (Coluna índice 18)
query = f"SELECT * FROM st_read('{caminho}', layer='RESUMO DIARIO') LIMIT 6 OFFSET 11"
dados = con.execute(query).fetchall()

# Função seca pra limpar o dado ou virar zero se vier lixo
def limpar(valor):
    try:
        return float(valor) if valor is not None else 0.0
    except:
        return 0.0

# Atribuindo cada linha pra sua variável
v_011 = limpar(dados[0][18])  # Linha 12
v_201 = limpar(dados[1][18])  # Linha 13
v_204 = limpar(dados[2][18])  # Linha 14
v_207 = limpar(dados[3][18])  # Linha 15
v_208 = limpar(dados[4][18])  # Linha 16
v_209 = limpar(dados[5][18])  # Linha 17

print(f"011: {v_011} |\n201: {v_201} |\n204: {v_204}")
print(f"207: {v_207} |\n208: {v_208} |\n209: {v_209}")