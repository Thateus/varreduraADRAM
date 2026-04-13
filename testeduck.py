import duckdb

con = duckdb.connect('boletimdiario.db')

con.execute("INSTALL spatial; LOAD spatial;")

caminho = "D:/pcp.faxinal/Documentos/py/BOLETIM DIARIO 31.xlsx"

con.execute(f"CREATE TABLE IF NOT EXISTS consolidado AS SELECT * FROM st_read('{caminho}')")

print(con.execute("SELECT count(*) FROM consolidado").fetchall())