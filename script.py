#Compara dois arquivos xlsx de e-mails e mostra quem está em uma, mas não está na outra.

import pandas as pd

# carregar arquivos Excel
members = pd.read_excel("members.xlsx") #arquivo baixado do e-mail serti (admin) para conferir se todos os usuários estão na lista do thiago
aniversariantes = pd.read_excel("aniversariantes.xlsx") #arquivo enviado pelo Thiago

# MOSTRAR COLUNAS (pra conferir nomes)
print("Colunas MEMBERS:", members.columns)
print("Colunas ANIVERSARIANTES:", aniversariantes.columns)

# 👉 AJUSTA AQUI SE PRECISAR
col_members = "Member Email" #nome das colunas nos arquivos que é feira a comparação
col_aniv = "E-mail para Contato"

# padronizar (minúsculo + remover espaços)
lista1 = set(members[col_members].astype(str).str.strip().str.lower())
lista2 = set(aniversariantes[col_aniv].astype(str).str.strip().str.lower())

# comparar
nao_encontrados = lista2 - lista1

print("\nUsuários que NÃO estão na lista principal:\n")

for email in sorted(nao_encontrados):
    print(email)

print(f"\nTotal: {len(nao_encontrados)}")

# 💾 salvar resultado em arquivo
resultado = pd.DataFrame({"Email": list(nao_encontrados)})
resultado.to_excel("nao_encontrados.xlsx", index=False)

print("\nArquivo 'nao_encontrados.xlsx' gerado com sucesso!")