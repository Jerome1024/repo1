# 把本文件夹中的一个word文档导出为图片

import os
import platform
import subprocess
from PIL import Image
from pdf2image import convert_from_path
from docx import Document
from docx2pdf import convert



# 获得word文件的名字，返回一个word文档的列表
def get_docs_name():
    files_name = os.listdir()
    docs_list = []
    for doc in files_name:
        if doc.endswith(".doc") or doc.endswith(".docx"):
            docs_list.append(doc)

    return docs_list


# 把一个word文件转化成pdf
def word2pdf(file_name):
    return convert(file_name)
    # os_name = platform.system()
    # if os_name == "Windows":
        # return convert(file_name)
    
    # elif os_name == "Linux":

    #     subprocess.run(['lowriter', '--headless', '--convert-to', 'pdf', file_name])
    #     return file_name.rsplit('.', 1)[0] + '.pdf'

    # else:
    #     print("不支持该操作系统")
    #     return None


# 获得一个word文档的页边距,返回一个页边距的字典
def get_margin(doc_path):
    doc = Document(doc_path)
    section = doc.sections[0]
    margins = {
            'top': section.top_margin.cm,
            'bottom': section.bottom_margin.cm,
            'left': section.left_margin.cm,
            'right': section.right_margin.cm
        }

    return margins


#把pdf 每一页转换成一个图片，返回的是一套图片的列表
def pdf2images(file_name):
    return convert_from_path(file_name)


#裁减一个图像列表中的每一个图片成需要的范围,参数是一个图片列表，一个页边距的字典
def crop_images(imags_list,margins, dpi=72):
    cropped_images= []

    for image in imags_list:
        width, height = image.size
        # left = margins['left'] * dpi
        left = 0
        top = margins['top'] * dpi
        # right = width - margins['right'] * dpi
        right = width
        bottom = height - margins['bottom'] * dpi

        cropped_image = image.crop((left,top,right,bottom))
        cropped_images.append(cropped_image)

    return cropped_images


#把一个图像列表中的图像合并成一张图片
def images_to_one(images):
    width, height = images[0].size
    combined_image = Image.new('RGB', (width, height * len(images)), color='white')  
    for i, image in enumerate(images):
        combined_image.paste(image, (0, i*height))

    return combined_image


# 调整美化最后的图片
def add_white_border(image, border,dpi=72):
    width, height = image.size
    
    new_height = height + 2*border*dpi
    new_image = Image.new('RGB', (width, new_height), 'white')
    new_image.paste(image, (0,border*dpi))

    return new_image


if __name__ == '__main__':

    docs_name = get_docs_name()
    current_dir = os.getcwd()

    for doc in docs_name:
        pdf_name = word2pdf(doc)
        
        if pdf_name:
            imag_list = pdf2images(pdf_name)
            margins = get_margin(doc)
            crop_imag_list = crop_images(imag_list, margins)
            #整合图片为一张长图，并且保存图片
            one_pic = images_to_one(crop_imag_list)
            re_pic = add_white_border(one_pic, int(margins['top']))
            if re_pic:
                pic_name = doc.rsplit('.', 1)[0] + '.png'
                output_path = os.path.join(current_dir, pic_name)
                re_pic.save(output_path)
        


    print("=======Finished=======")
