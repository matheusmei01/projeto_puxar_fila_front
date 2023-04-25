import pytesseract
import cv2

imagem = cv2.imread("email.JPG")

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'
texto = pytesseract.image_to_string(imagem) 
print(texto)