from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from PIL import Image

#単品商品のpdfの作成
def write_single_item_pdf(mail_id, item_name_list, delivery_method, delivery_time, receipt_num, item_password, item_description_list):
    pdf = canvas.Canvas("pdfs\\{}.pdf".format(mail_id))
    pdf.setPageSize((62*mm, 100*mm))
    png = Image.open("QRcode\\{}.png".format(mail_id))

    item_name_len = len(item_name_list)
    pdf.drawInlineImage(png, 15*mm, 68*mm, 30*mm, 30*mm)
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    pdf.setFont('HeiseiKakuGo-W5', 8)

    for i, item_name in enumerate(item_name_list):
        pdf.drawString(5*mm, (45 + i*5)*mm, item_name)

    pdf.drawString(5*mm, 35*mm, '配送方法 :   {}'.format(delivery_method))
    pdf.drawString(5*mm, 30*mm, '配送希望時間帯 :   {}'.format(delivery_time))
    pdf.drawString(5*mm, 25*mm, '受付番号   :   {}'.format(receipt_num))
    pdf.drawString(5*mm, 20*mm, 'パスワード :   {}'.format(item_password))
    pdf.drawString(5*mm, 15*mm, 'Product_ID    :   {}'.format(item_description_list[0]))
    pdf.drawString(5*mm, 10*mm, 'Quantity   :   {}'.format(item_description_list[1]))
    pdf.drawString(5*mm, 5*mm, 'Order_No    :   {}'.format(item_description_list[2]))

    pdf.save()


#まとめ取り引きのpdf作成
def write_group_item_pdf(mail_id, delivery_method, delivery_time, receipt_num, item_password, items_description_list):

    # まとめ取引のpdfを作成するときの配置を指定する数字
    a =40*(len(items_description_list))

    pdf = canvas.Canvas("pdfs\\まとめて取引_{}.pdf".format(mail_id))
    pdf.setPageSize((62*mm, (a + 60)*mm))
    png = Image.open("QRcode\\まとめて取引_{}.png".format(mail_id))

    pdf.drawInlineImage(png, 15*mm, (a + 28)*mm, 30*mm, 30*mm)
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    pdf.setFont('HeiseiKakuGo-W5', 8)

    pdf.drawString(5*mm, (a + 20)*mm, '配送方法 :   {}'.format(delivery_method))
    pdf.drawString(5*mm, (a + 15)*mm, '配送希望時間帯 :   {}'.format(delivery_time))
    pdf.drawString(5*mm, (a + 10)*mm, '受付番号   :   {}'.format(receipt_num))
    pdf.drawString(5*mm, (a + 5)*mm, 'パスワード :   {}'.format(item_password))
    pdf.drawString(5*mm, (a + 0)*mm, '---------------------------------------------')

    for j, item_description_list in enumerate(items_description_list):
        for i, item_name in enumerate(item_description_list[0]):
            pdf.drawString(5*mm, ((20 + i*5) + 40*j)*mm, item_name)

        pdf.drawString(5*mm, (15 + 40*j)*mm, 'Product_ID    :   {}'.format(item_description_list[1]))
        pdf.drawString(5*mm, (10 + 40*j)*mm, 'Quantity   :   {}'.format(item_description_list[2]))
        pdf.drawString(5*mm, (5 + 40*j)*mm, 'Order_No    :   {}'.format(item_description_list[3]))
        pdf.drawString(5*mm, (0 + 40*j)*mm, '---------------------------------------------')

    pdf.save()
