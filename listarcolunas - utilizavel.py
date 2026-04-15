import duckdb

con = duckdb.connect()
con.execute("INSTALL spatial")
con.execute("LOAD spatial")

caminho = r"D:\pcp.faxinal\Documentos\py\BOLETIM DIARIO 31.xlsx"

query = """
SELECT *
FROM st_read(?, layer='RESUMO DIARIO')
LIMIT 6 OFFSET 11
"""

dados = con.execute(query, [caminho]).fetchall()

def limpar(valor):
    try:
        return float(valor)
    except:
        return 0.0

chaves = ["011", "201", "204", "207", "208", "209"]

valores = {k: limpar(dados[i][18]) for i, k in enumerate(chaves)}

total = sum(valores.values())

print("\n===RESUMO DIARIO===\n")
for k, v in valores.items():
	status = "OK" if v > 0 else "ZERO"
	print(f"Produto COD. {k}: {v:.0f} [{status}]")

print(f"\nTOTAL: {total:.0f}")