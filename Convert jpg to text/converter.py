from PIL import Image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
import os

# Указываем путь к изображению
image_path = os.path.abspath(input())

# Открываем изображение с помощью Pillow
image = Image.open(image_path)

# Используем Pytesseract для распознавания текста
text = pytesseract.image_to_string(image, lang='rus')

# Выводим распознанный текст
print(text)