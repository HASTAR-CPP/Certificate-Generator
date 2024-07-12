from PIL import Image, ImageDraw, ImageFont
import pandas as pd
form = pd.read_excel("Certificate_PCB.xlsx")
name_list = form['Name'].str.upper().to_list()
company_list = form['Company'].to_list()
signature_path = "/home/kali/Documents/Programming/python_project1/signature.png"
output_path = '/home/kali/Documents/Programming/python_project1/certificate/'
for i,j in zip(name_list,company_list):
    im = Image.open("certificate-of-participation.png").convert("RGB")  
    # im = Image.open("certificate-of-participation.png")
    d = ImageDraw.Draw(im)
    # writing name
    location_name = (300, 640)
    text_color = (0, 35, 255)
    font = ImageFont.truetype("lcallig.ttf", 70)
    d.text(location_name, i, fill=text_color, font=font)
    # writing company
    location_company = (622,760)
    font = ImageFont.truetype("lcallig.ttf", 40)
    d.text(location_company, j, fill=text_color, font=font)
    # writing year after competition
    location_year = (1146,760)
    d.text(location_year, text="2024", fill=text_color, font=font)
    # writing date
    location_date = (189,860)
    d.text(location_date, text="12/07/2024", fill=text_color, font=font)
    # pasting signature
    signature = Image.open(signature_path).convert("RGBA")
    signature = signature.resize((200,100))
    signature_location = (1060,835)
    im.paste(signature,signature_location,signature)

    print("Saving - "+i)
    im.save(output_path + "/certificate_"+ i + ".pdf")