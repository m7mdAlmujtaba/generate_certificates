import numpy as np
import uuid
from PIL import ImageFont, ImageDraw, Image
import qrcode
import img2pdf
import xlrd
import xlwt
from xlutils.copy import copy


def im2pdf(name, course, from_date, to_date, serial, instructor, director, directory):
    # Upload The template
    img = Image.open('img/empty_certificate.png')
    w, h = img.size

    date = 'Date: ' + from_date +'To ' + to_date

    nc_font_path = "files/fonts/Proxima-Nova-Bold.otf"
    isd_font_path = "files/fonts/Poppins-SemiBold.ttf"

    student_id = uuid.uuid1()
    print(student_id)
    qr = qrcode.make(student_id)
    type(qr)  # qrcode.image.pil.PilImage
    
    basewidth = 300
    wpercent = (basewidth/float(qr.size[0]))
    hsize = int((float(qr.size[1])*float(wpercent)))
    qr = qr.resize((basewidth,hsize), Image.ANTIALIAS)

    img.paste(qr, (2395 - int(basewidth/2), 1470))

    name_font = ImageFont.truetype(nc_font_path, 72)
    course_font = ImageFont.truetype(isd_font_path, 60)
    date_font = ImageFont.truetype(isd_font_path, 34)
    serial_font = ImageFont.truetype(isd_font_path, 40)
    ins_font = ImageFont.truetype(isd_font_path, 30)
    drc_font = ImageFont.truetype(isd_font_path, 30)
    draw = ImageDraw.Draw(img)

    # To assure that the name will not get out of the certificate
    n_w, n_h = draw.textsize(name, font=name_font)
    print(name, n_w)
    x = 70
    while n_w > (w*.8):
        name_font = ImageFont.truetype(nc_font_path, x)
        n_w, n_h = draw.textsize(name, font=name_font)
        x = x - 1
    print(name, n_w)

    n_w, n_h = draw.textsize(name, font=name_font)
    c_w, c_h = draw.textsize(course, font=course_font)
    d_w, d_h = draw.textsize(date, font=date_font)
    s_w, s_h = draw.textsize(serial, font=serial_font)
    i_w, i_h = draw.textsize(instructor, font=ins_font)
    drc_w, drc_h = draw.textsize(director, font=ins_font)

    draw.text(((w-n_w)/2, (h-n_h*2)/2+30), name, font=name_font, fill="black")
    draw.text(((w-c_w)/2, 1250), course, font=course_font, fill=(6, 105, 176))
    draw.text(((w-d_w)/2, 1420), date, font=date_font, fill="black")
    draw.text((2395-s_w/2, 820), serial, font=serial_font, fill="white")
    draw.text((1700-i_w/2, 1670), instructor, font=ins_font, fill="black")
    draw.text((1000 - drc_w / 2, 1670), director, font=drc_font, fill="black")

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
    
    rb = xlrd.open_workbook('files/ids.xls')
    r_sheet = rb.sheet_by_index(0)
    r = r_sheet.nrows

    wb = copy(rb) 
    sheet = wb.get_sheet(0) 

    sheet.write(r, 0, serial[3:])
    sheet.write(r, 1, str(student_id))

    wb.save('files/ids.xls')
    # output
    print("Successfully made pdf file")


#im2pdf('Mohammed Almujtaba Ali Hassan Musa Mohammed Almujtaba Ali Hassan Musa', 'Javascript Advanced', '1/2/2022', '1/3/2022', 'SN:20202020', 'Omar Abdallah', 'Amjad', 'pdf/')