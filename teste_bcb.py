from bcb import sgs

print("Conectando ao Banco Central...")

try:
    # Tenta puxar apenas os últimos 5 dias da Selic para ser rápido
    selic = sgs.get({'SELIC': 4390}, last=5)
    print("✅ Sucesso! O BCB respondeu. Dados recebidos:")
    print(selic)
except Exception as e:
    print("❌ FALHA NA CONEXÃO. O erro real enviado pelo sistema foi:")
    print(repr(e))