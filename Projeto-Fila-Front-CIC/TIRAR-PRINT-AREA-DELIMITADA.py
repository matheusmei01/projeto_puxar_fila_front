from PIL import ImageGrab

# Capturar uma área específica da tela

left = 1206
top = 672
right = 1241
bottom = 694

imagem = ImageGrab.grab((left, top, right, bottom))

# Exibir a imagem capturada
imagem.show()