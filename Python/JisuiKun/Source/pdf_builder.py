from PIL import Image
import os

def build_pdf(save, files, cover=None):
    images=[]
    if cover:
        images.append(Image.open(cover).convert("RGB"))

    for f in files:
        images.append(Image.open(os.path.join(save,f)).convert("RGB"))

    if images:
        images[0].save(os.path.join(save,"cap_output.pdf"),
                       save_all=True,append_images=images[1:])
