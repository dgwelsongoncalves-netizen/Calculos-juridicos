from bcb import sgs

print("Buscando IPCA e Selic no Banco Central...")
# Puxa IPCA (433) e Selic (11) simultaneamente
dados = sgs.get({'IPCA': 433, 'SELIC': 11}, last=5)

print("\nÚltimos índices consolidados:")
print(dados)