from PIL import ImageGrab
from PIL import Image
import pytesseract
import cv2

# Capturar uma área específica da tela
left = 1209
top = 677
right = 1240
bottom = 693
imagem = ImageGrab.grab((left, top, right, bottom))

# Exibir a imagem capturada
imagem.show()

imagem.save("imagem.png")

# Abrir a imagem
novaimagem = cv2.imread("imagem.png")

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'
texto = pytesseract.image_to_string(imagem) 
print(texto)