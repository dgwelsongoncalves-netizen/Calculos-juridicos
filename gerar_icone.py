from PIL import Image

# Coloque o nome da sua imagem original gigante de 6 MB aqui (pode ser .png ou .jpg)
caminho_imagem_original = "Nash.png" 

# Abre a imagem de alta resolução
img = Image.open(caminho_imagem_original)

# Empacota em todas as resoluções que o Windows precisa
tamanhos = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

# Salva o arquivo .ico profissional
img.save("dr_nash.ico", sizes=tamanhos)

print("Ícone de alta resolução gerado com sucesso!")