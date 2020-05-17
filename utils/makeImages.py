from PIL import Image, ImageDraw, ImageFont

def make_image(text):
    img_size = (300, 300)
    text_size = 100
    text_anchor = ((img_size[0]-text_size)/2, (img_size[1]-text_size)/2)
    im = Image.new("RGB", img_size, "black")# Imageインスタンスを作る
    draw = ImageDraw.Draw(im)# im上のImageDrawインスタンスを作る
    fnt = ImageFont.truetype("utils/font/msgothic.ttc", text_size) #ImageFontインスタンスを作る
    draw.text(text_anchor, text, font=fnt) #fontを指定
    im.save("./image/{0}.png".format(text))

make_image("１")
make_image("２")
make_image("３")
make_image("４")
make_image("５")
make_image("６")
make_image("７")
make_image("８")
make_image("９")

