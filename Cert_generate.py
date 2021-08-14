import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# read list contain participants name
data = pd.read_excel('Name_list/names.xlsx')
# fetch and store name column values 
name_list = data["Name"].tolist()
# certificate model resolution
W,H = (2000,1414)
# text color of name 
text_color = (0, 0, 0)
# font of name
font = ImageFont.truetype("Fonts/Oswald-Bold.ttf", 50)

#generate certificate
print("certificate generation START")
for name in name_list:
    image = Image.open("Model_Certificate/cert_model.png")      
    background = Image.new("RGB", image.size, (255, 255, 255))
    # 3 is the alpha channel
    background.paste(image, mask=image.split()[3])                 
    image=background
    d = ImageDraw.Draw(image)
    # text size
    w,h = d.textsize(name)
    # adjust name to center 
    location = ((W-w)/2 -w+10,(H-h)/2 -30)
    # add name
    d.text(location, name, fill = text_color, font = font)
    # save certificates in pdf format
    image.save("Generate/CID_" + name + ".pdf")
print("certificate generation END")
