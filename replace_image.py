from docxtpl import DocxTemplate
import io

DEST_FILE = 'output/header_footer_image.docx'

tpl=DocxTemplate('templates/header_footer_image_tpl.docx')

context = {
    'mycompany' : 'The World Wide company',
}
# 根据图片占位，替换图片 这种替换可以提前在模版中预设图片大小
django_pic = io.BytesIO(open('templates/django.png', 'rb').read())
dummy_pic = io.BytesIO(open('templates/dummy_pic_for_header.png', 'rb').read())
python_image = io.BytesIO(open('templates/python.png', 'rb').read())
tpl.replace_media(dummy_pic, django_pic)
tpl.replace_media(django_pic, python_image)

tpl.render(context)
tpl.save(DEST_FILE)