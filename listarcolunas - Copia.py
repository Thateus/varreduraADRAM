import duckdb

con = duckdb.connect('boletim.db')
con.execute("INSTALL spatial; LOAD spatial;")

caminho = r"D:\pcp.faxinal\Documentos\py\BOLETIM DIARIO 31.xlsx"

# LIMIT 6 pega as 6 linhas (12, 13, 14, 15, 16, 17)
# OFFSET 11 pula as primeiras 11 e começa na 12
query = f"SELECT * FROM st_read('{caminho}', layer='RESUMO DIARIO') LIMIT 6 OFFSET 11"
bloco_de_linhas = con.execute(query).fetchall()

print("--- EXTRAÇÃO DA COLUNA P (Índice 18) ---")

valores_p = []

for i, linha in enumerate(bloco_de_linhas):
    num_linha_excel = 12 + i
    valor_bruto = linha[18]
    
    # TRATAMENTO DE CHOQUE:
    try:
        # Tenta converter para número (float aceita decimais e inteiros)
        valor_limpo = float(valor_bruto) if valor_bruto is not None else 0.0
    except (ValueError, TypeError):
        # Se for texto que não vira número (ex: "OBS"), vira zero
        valor_limpo = 0.0
        
    valores_p.append(valor_limpo)
    print(f"Célula P{num_linha_excel}: {valor_bruto} -> (Tratado como: {valor_limpo})")
