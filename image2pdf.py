import numpy as np
import uuid
from PIL import ImageFont, ImageDraw, Image
import qrcode
import img2pdf
import xlrd
import xlwt
from xlutils.copy import copy


def im2pdf(name, course, date, serial, instructor):
    # Upload The template
    img = Image.open('img/empty_certificate.png')
    w, h = img.size

    # date = 'Date: 5-Aug-2018'

    nc_font_path = "BASKVILL.TTF"
    isd_font_path = "BebasKai.otf"

    student_id = uuid.uuid1()
    print(student_id)
    qr = qrcode.make(student_id)
    type(qr)  # qrcode.image.pil.PilImage
    
    basewidth = 200
    wpercent = (basewidth/float(qr.size[0]))
    hsize = int((float(qr.size[1])*float(wpercent)))
    qr = qr.resize((basewidth,hsize), Image.ANTIALIAS)

    img.paste(qr, (1530 - int(basewidth/2), 940))

    name_font = ImageFont.truetype(nc_font_path, 60)
    course_font = ImageFont.truetype(nc_font_path, 50)
    date_font = ImageFont.truetype(isd_font_path, 25)
    serial_font = ImageFont.truetype(isd_font_path, 45)
    ins_font = ImageFont.truetype(isd_font_path, 25)
    drc_font = ImageFont.truetype(isd_font_path, 25)
    draw = ImageDraw.Draw(img)

    n_w, n_h = draw.textsize(name, font=name_font)
    c_w, c_h = draw.textsize(course, font=course_font)
    d_w, d_h = draw.textsize(date, font=date_font)
    s_w, s_h = draw.textsize(serial, font=serial_font)
    i_w, i_h = draw.textsize(instructor, font=ins_font)
    drc_w, drc_h = draw.textsize('MOHAMMED ALMUJTABA', font=ins_font)

    draw.text(((w-n_w)/2, (h-n_h*2)/2+15), name, font=name_font, fill="black")
    draw.text(((w-c_w)/2, 775), course, font=course_font, fill="black")
    draw.text(((w-d_w)/2, 920), date, font=date_font, fill=(142, 189, 143, 0))
    draw.text((1530-s_w/2, h/2 - 125), serial, font=serial_font, fill="white")
    draw.text((1150-i_w/2, 1030), instructor, font=ins_font, fill="black")
    draw.text((705 - drc_w / 2, 1030), "MOHAMMED ALMUJTABA", font=drc_font, fill="black")

    # img.show()

    # To PDF
    img = np.array(img)
    image_without_alpha = img[:,:,:3]
    img = Image.fromarray(image_without_alpha)
    img.save('img/certi.png')
    img.close()
    img = Image.open('img/certi.png')

    pdf_path = "pdf/"+serial[3:]+".pdf"

    # converting into chunks using img2pdf
    pdf_bytes = img2pdf.convert(img.filename)

    # opening or creating pdf file
    file = open(pdf_path, "wb")

    # writing pdf files with chunks
    file.write(pdf_bytes)

    # closing image file
    img.close()

    # closing pdf file
    file.close()
    
    rb = xlrd.open_workbook('ids.xls')
    r_sheet = rb.sheet_by_index(0)
    r = r_sheet.nrows + 1 

    wb = copy(rb) 
    sheet = wb.get_sheet(0) 

    sheet.write(r, 0, serial[3:])
    sheet.write(r, 1, str(student_id))

    wb.save('ids.xls')
    # output
    print("Successfully made pdf file")


