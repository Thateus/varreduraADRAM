import duckdb

con = duckdb.connect()
con.execute("LOAD spatial")

caminho = r"\\ARAGORN\Fabrica\PCP - Planejamento e Controle de Produção\BOLETIM DIARIO - MATHEUS\BOLETIM 2026\MARÇO\BOLETIM DIARIO 31.xlsx"

query = """
SELECT *
FROM st_read(?, layer='Produção Diária')
"""

dados = con.execute(query, [caminho]).fetchall()

def limpar(valor):
    try:
        return float(valor)
    except:
        return 0.0

COL_CODIGO = 1
COL_DESCRICAO = 2
COL_VALOR = 3

chaves = ["011", "201", "204", "207", "208", "209"]

valores = {k: {"total": 0, "nome": ""} for k in chaves}

for linha in dados:
	codigo = str(linha[COL_CODIGO]).strip()

	if not codigo.isdigit():
		continue

	if codigo in valores:
		valores[codigo]["total"] += limpar(linha[COL_VALOR])

		nome = str(linha[COL_DESCRICAO]).strip() 
		if nome:
			valores[codigo]["nome"] = nome

total_geral = sum(v["total"] for v in valores.values())

print("\n=== PRODUÇÃO MENSAL===\n")

for k, info in valores.items():
	nome = info["nome"] or "SEM NOME"
	total = int(info["total"])
	status = "OK" if total > 0 else "ZERO"

	print(f"{nome} ({k}): {total} [{status}]")

print(f"\nTOTAL: {total_geral:.0f}\n\n")