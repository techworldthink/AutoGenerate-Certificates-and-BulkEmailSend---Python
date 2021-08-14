import pandas as pd
from PIL import Image, ImageDraw, ImageFont

data = pd.read_excel('Name_list/names.xlsx')
name_list = data["Name"].tolist()
print(name_list)

for i in name_list:
    im = Image.open("Model_Certificate/cert_model.png")
    #im.show()
    background = Image.new("RGB", im.size, (255, 255, 255))
    background.paste(im, mask=im.split()[3]) # 3 is the alpha channel
    #background.show()
    im=background
    d = ImageDraw.Draw(im)
    location = (100, 398)
    text_color = (0, 137, 209)
    font = ImageFont.truetype("Fonts/Oswald-Bold.ttf", 120)
    d.text(location, i, fill = text_color, font = font)
    im.save("Generate/CID_" + i + ".pdf")
