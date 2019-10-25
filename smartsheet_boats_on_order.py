#!/usr/bin/env python3
"""
convert smartsheet boats on order reports to specially formatted execl sheets and pdfs

   NOTES: page lenght for all pages
            is_clemens()  'p navs to sets the page length for the FIRST PAGE
          process_rows()  'q navs to the first IF sets the page length for remaining pages 

"""

import smartsheet
# import logging
import datetime
import glob
import os
import subprocess
import openpyxl
import dateparser
import datedelta
import click
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from emailer import *
from PyPDF2 import PdfFileReader, PdfFileWriter

api = os.getenv('SMARTSHEET_API')
source_dir = os.getenv('SOURCE_DIR')
target_dir = os.getenv('TARGET_DIR')
rollover = int(os.getenv('ROLLOVER'))
one_date_fmt = os.getenv('ONEDATEFMT')
two_date_fmt = os.getenv('TWODATEFMT')

full_month = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
short_month = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']


reports = [
    {
        'id': 6215848202397572,
        'name': 'Alaska Frontier Fabrication - Boats on Order'
    },
    {
        'id': 1862555250517892,
        'name': 'Boat Country - Boats on Order'
    },
    {
        'id': 3611603372402564,
        'name': 'Clemens Eugene - Boats on Order'
    },
    {
        'id': 7685431392266116,
        'name': 'Clemens Portland - Boats on Order'
    },
    {
        'id': 6603151173281668,
        'name': 'Elephant Boys - Boats on Order'
    },
    {
        'id': 8501389329491844,
        'name': 'Enns Brothers - Boats on Order'
    },
    {
        'id': 7533698787633028,
        'name': 'Idaho Marine - Boats on Order'
    },
    {
        'id': 5291433801344900,
        'name': 'PGM - Boats on Order'
    },
    {
        'id': 3591949602056068,
        'name': 'Port Boat House - Boats on Order'
    },
    {
        'id': 7351798332712836,
        'name': 'RF Marina - Boats on Order'
    },
    {
        'id': 4536017773455236,
        'name': 'The Bay Co - Boats on Order'
    },
    {
        'id': 7159452517328772,
        'name': 'Three Rivers - Boats on Order'
    },
    {
        'id': 875382787336068,
        'name': 'Valley Marine - Boats on Order'
    },
    {
        'id': 7940135837820804,
        'name': 'Y Marina - Boats on Order'
    },
    {
            'id': 4374979853739908,
            'name': 'Avataa - Boats on Order'
    },
    {
            'id': 2931819302676356,
            'name': 'All Dealer - Boats on Order'
    },
]


log_text = ""
errors = False
titles = []
order_details = 0 
row_offset = 0
max_column = 0
template_file = ''
clemens = 0
page_number = 0
page_breaks_normal = [
  50,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
  56,
]

page_breaks_clemens = [
  58,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
  62,
]

def log(text, error=None):
    global log_text, errors
    print(text)
    log_text += text + "\n"
    if (error):
        errors = True

def mail_results(subject, body):
    mFrom = os.getenv('MAIL_FROM')
    mTo = os.getenv('MAIL_TO')
    m = Email(os.getenv('MAIL_SERVER'))
    m.setFrom(mFrom)
    m.addRecipient(mTo)
    m.addCC(os.getenv('MAIL_ALSO'))
   
    m.setSubject(subject)
    m.setTextBody("You should not see this text in a MIME aware reader")
    m.setHtmlBody('<pre>\n' + body + '</pre>\n')
    m.send()


class Dealership:
    pass

class Dealerships:
    pass

class Column:
    name = 'Arial'
    size = '9'
    bold = False
    italic = False
    color = '000000'
    bg_color = 'FFFFFF'

    def __init__(self, old, new):
        self.text = ''
        self.old_column = old 
        self.new_column = new
        self.reset()

    def reset (self):
        self.text = ''
        self.name = ''
        self.size = ''
        self.bold = False 
        self.italic = False
        self.color = ''
        self.bg_color = ''

    def font (self):
        return Font(
            name=self.face or Column.face,
            size=self.size or Column.size,
            bold=self.bold or Column.bold,
            itailc=self.italic or Column.italic,
            color=self.color or Column.color
        )


def adjustDate(myDate):
    if myDate.day >= rollover:
        while myDate.day > 1:
            myDate = myDate + datedelta.DAY
    return myDate

def startInfo(value):
    output = ''
    textColor = '000000'
    dates = value.split('/')

    # process start date
    start = dateparser.parse(dates[0], settings={'PREFER_DATES_FROM': 'future'})

    # check for null start date
    if not start:
        return [output, textColor]

    # round up to the next month
    startDate = adjustDate(start)

    # Set colors for this month or next month
    if startDate.month == datetime.date.today().month:
        textColor = 'B00000'
    elif startDate.month == (datetime.date.today() + datedelta.MONTH).month:
        textColor = '0000F0'
    # set output in case we are only outputting a start date
    output = startDate.strftime(one_date_fmt)

    # no end date
    if len(dates) == 1:
        return [output, textColor]

    # process end date
    end = dateparser.parse(dates[1], settings={'PREFER_DATES_FROM': 'future'})
    
    # check for null end date
    if not end:
        return [output, textColor]

    endDate = adjustDate(end)
    output = startDate.strftime(two_date_fmt) + ' / ' + endDate.strftime(two_date_fmt)
    return [output, textColor]

def is_clemens(flag):
    global order_details, row_offset, template_file, max_column, titles, clemens
    if flag:
        clemens = 1
        template_file = source_dir + 'BoatsOnOrderTemplateClemens.xlsx'
        order_details = 6 
        row_offset = 1
        max_column = 12
        titles = ['Hull #',
                  'Boat Model',
                  'Order Details',
                  'Colors Interior / Exterior',
                  'Engines',
                  'Order Details',
                  'Current Phase',
                  'Est Start/Finish',
                  'Actual Start',
                  'Actual Finish',
                  'Notes']
    else:
        clemens = 0
        template_file = source_dir + 'BoatsOnOrderTemplate.xlsx'
        order_details = 3
        row_offset = 0
        max_column = 11
        titles = ['Hull #',
                  'Boat Model',
                  'Order Details',
                  'Colors Interior / Exterior',
                  'Engines',
                  'Current Phase',
                  'Est Start/Finish',
                  'Notes',
                  '',
                  '']

def normal_border(wsNew, row):
    for i in range(1, max_column - 2):
        side1 = 'thin'
        side2 = 'thin'
        if i == max_column - 3:
            side1 = 'medium'
        if i == 1:
            side2 = 'medium'
        wsNew.cell(column=i, row=row+7).border = Border(
            right=Side(border_style=side1, color='FF000000'),
            left=Side(border_style=side2, color='FF000000'))


def side_border(wsNew, row):
    wsNew.cell(column=1, row=row+7).border = Border(
        left=Side(border_style='medium', color='FF000000'))
    wsNew.cell(column=max_column - 3, row=row+7).border = Border(
        right=Side(border_style='medium', color='FF000000'))


def heading_border(wsNew, row):
    for i in range(1, max_column-2):
        side1 = 'thin'
        side2 = 'thin'
        if i == max_column - 1:
            side1 = 'medium'
        if i == 1:
            side2 = 'medium'
        wsNew.cell(column=i, row=row+7).border = Border(
            right=Side(border_style=side1, color='FF000000'),
            left=Side(border_style=side2, color='FF000000'),
            top=Side(border_style='medium', color='FF000000'),
            bottom=Side(border_style='medium', color='FF000000'))


def end_page_border(wsNew, row):
    for i in range(1, max_column -2):
        side1 = 'thin'
        side2 = 'thin'
        if i == max_column - 3:
            side1 = 'medium'
        if i == 1:
            side2 = 'medium'
        wsNew.cell(column=i, row=row+7).border = Border(
            right=Side(border_style=side1, color='FF000000'),
            left=Side(border_style=side2, color='FF000000'),
            bottom=Side(border_style='medium', color='FF000000'))


def bottom_border(wsNew, row):
    for i in range(1, max_column - 2):
        side1 = 'thin'
        side2 = 'thin'
        if i == max_column - 3:
            side1 = 'medium'
        if i == 1:
            side2 = 'medium'
        wsNew.cell(column=i, row=row+7).border = Border(
            right=Side(border_style=side1, color='FF000000'),
            left=Side(border_style=side2, color='FF000000'),
            bottom=Side(border_style='medium', color='FF000000'))

def fetch_value(cell):
    value = cell.value
    if cell.data_type == 's':
        return value
    if cell.is_date:
        return ('%02d/%02d/%02d' % (
            value.month,
            value.day,
            value.year - 2000))
    if value is None:
        return ''
    return str(int(value))

def get_font_color(start_date):
    if (start_date < datetime.date.today() + datetime.timedelta(days=15)):
        return 'B00000'
    elif (start_date < datetime.date.today() + datetime.timedelta(days=29)):
        return '0000F0'
    else:
        return '000000' 

def set_mast_header(wsNew, logo_name, dealer_name):
    # place logo and dealername on new sheet
    img = Image(logo_name)
    wsNew.add_image(img, 'B1')
    wsNew['B5'] = dealer_name
    wsNew['H5'] = "Report Date: %s " % (
        datetime.datetime.today().strftime('%m/%d/%Y'))


def set_header(wsNew, row):
    heading_border(wsNew, row)
    wsNew.row_dimensions[row+7].height = 21.6

    for i in range(1, max_column - 2):
        wsNew.cell(row=row+7, column=i, value=titles[i-1])
        wsNew.cell(row=row+7, column=i).alignment = Alignment(
            horizontal='center', vertical='center')
        wsNew.cell(row=row+7, column=i).font = Font(bold=True, size=9, name='Arial')


def set_footer(wsNew, row):
    side_border(wsNew, row)
    side_border(wsNew, row+1)

    wsNew.merge_cells(start_row=row+8,
                      start_column=1,
                      end_row=row+8,
                      end_column=3)
    wsNew.cell(row=row+8, column=1,
               value="Contact Joe for 9'6 build dates")
    wsNew.cell(row=row+8, column=1).alignment = Alignment(horizontal='center')
    wsNew.cell(row=row+8, column=1).font = Font(bold=True)

    wsNew.merge_cells(start_row=row+9,
                      start_column=1,
                      end_row=row+9,
                      end_column= max_column - 3)
    wsNew.cell(row=row+9, column=1,
               value=("NOTE: Estimated Start & Delivery Week's can be 1 - 2 "
                      "Weeks before or after original dates"))
    wsNew.cell(row=row+9, column=1).alignment = Alignment(horizontal='center')
    wsNew.cell(row=row+9, column=1).font = Font(bold=True)
    bottom_border(wsNew, row+2)


def process_row(wsOld, wsNew, row, offset, bgColor, base):  # base 7 or base 0
    global row_offset, clemens
    split = 8 + clemens
    
    estimate = fetch_value(wsOld.cell(column=(7 + row_offset), row=row))
    dates, font_color = startInfo(estimate)

    for i in range(1, max_column - 2):
        skipcell = i if i < split else i + 2
        value = fetch_value(wsOld.cell(column=skipcell, row=row))
        if (i == 7 and clemens == 0) or (i == 8 and clemens == 1):
           value = dates

        cell = wsNew.cell(column=i, row=row+base+offset)
        cell.value = value
        bg = bgColor
        if i == order_details and cell.value.lower().find('stock') == -1:
            bg = 'FFFFC000'
        if wsOld.cell(column=i, row=row).fill.start_color.index == 'FF00CA0E':
            bg = 'FF00CA0E'
        if bg is not None:
            cell.fill = PatternFill(start_color=bg,
                                    end_color=bg,
                                    fill_type="solid")
        if i == 2 or i == 4 + row_offset  or i == 5 + row_offset or i == max_column - 1 + row_offset:
            cell.font = Font(name='Arial',size=8, color=font_color)
        else:
            cell.font = Font(name='Arial',size=9, color=font_color)


def set_background_color(wsOld, row):
    """
    Change background color on OS and HardTops
    """
    model = wsOld["B"+str(row)].value
    if model is None:
        return None
    if model.find("OS") != -1:
        return 'FFA6A6A6'
    if model.replace(" ", "").lower().find('hardtop') != -1:
         return 'FFD9D9D9'
    return None


def process_rows(wsOld, wsNew, base, forPDF):
    global page_number 
    pagelen = page_breaks_clemens[page_number] if clemens else page_breaks_normal[page_number] 
    offset = 0
    last_page_offset = 0
    i = 4 

    for i in range(2, wsOld.max_row + 1):
        bgColor = set_background_color(wsOld, i)
        process_row(wsOld, wsNew, i, offset, bgColor, base)

        # if there are not 3 lines left for footer on last page handle
        if (i > wsOld.max_row - 4) and (i> pagelen - base - 3):
            last_page_offset = 3
            pagelen = i + base

        if (i + base) % pagelen == 0  and wsOld.max_row != i and forPDF:
            end_page_border(wsNew, i + offset)
            offset += 1 + last_page_offset

            if i < wsOld.max_row: # - (page_breaks_clemens[page_number] if clemens else page_breaks_normal[page_number]) + 1 :
                offset += 1
                set_header(wsNew, i + offset )

            page_number += 1
            pagelen += page_breaks_clemens[page_number] if clemens else page_breaks_normal[page_number]
        else:
            normal_border(wsNew, i + offset)

    end_page_border(wsNew, i + offset)
    offset += 1
    set_footer(wsNew, wsOld.max_row + offset)



def process_sheet_to_pdf(file):
    # change variables here
    input_name = source_dir + 'downloads/' + file
    watermark_name = source_dir + 'watermark.pdf'
    temp_name = source_dir + 'temp.xlsx'
    pdf_dir = (target_dir + 'Formatted - PDF/')
    output_name = pdf_dir + file
    logo_name = source_dir + 'nrblogo1.jpg'
    dealer_name = file[:-22]
    base = 7

    # load sheet data is coming from
    wbOld = openpyxl.load_workbook(input_name)
    wsOld = wbOld.active

    # load sheet we are copying data to
    wbNew = openpyxl.load_workbook(template_file)
    wsNew = wbNew.active

    set_mast_header(wsNew, logo_name, dealer_name)
    process_rows(wsOld, wsNew, base, True)

    range = 'A1:J'+str(wsNew.max_row + 10)

    wbNew.create_named_range('_xlnm.Print_Area', wsNew, range, scope=0)

    # save new sheet out to temp.xls and temp.pdf file
    try:
        wbNew.save(output_name)
        result = subprocess.call(['/usr/local/bin/unoconv',
                         '-f', 'pdf',
                         '-t', source_dir + 'landscape.ots',
                         '--output='+ temp_name[:-4] + 'pdf',
                         output_name])
        if (result):
            log('             UNICONV FAILED TO CREATE PDF', True)
    except Exception as e:
        log('             FAILED TO CREATE XLSX AND PDF: ' + str(e), True)
    # add watermark to temp.pdf and save to proper dealership name
    try:
        add_watermark(temp_name[:-4] + 'pdf', watermark_name, output_name[:-4] + 'pdf')
        # os.remove(temp_name[:-4] + 'pdf')
    except Exception as e:
        log('             FAILED TO ADD WATERMARK: ' + str(e), True)

     
def add_watermark(input, watermark, output):
    file = open(input, 'rb')
    reader = PdfFileReader(file)

    watermark = open(watermark,'rb')
    reader2 = PdfFileReader(watermark)
    waterpage = reader2.getPage(0)

    writer = PdfFileWriter()

    for pageNum in range(0, reader.numPages):
        page = reader.getPage(pageNum)
        page.mergePage(waterpage)
        writer.addPage(page)

    resultFile = open(output,'wb')
    writer.write(resultFile)
    file.close()
    resultFile.close()


def process_sheet_to_xlsx(file):
    # change variables here
    input_name = source_dir + 'downloads/' + file
    output_name = target_dir + file
    logo_name = source_dir + 'nrblogo1.jpg'
    dealer_name = file[:-22]
    base = 7

    # load sheet data is coming from
    wbOld = openpyxl.load_workbook(input_name)
    wsOld = wbOld.active

    # load sheet we are copying data to
    wbNew = openpyxl.load_workbook(template_file)
    wsNew = wbNew.active
    set_mast_header(wsNew, logo_name, dealer_name)
    process_rows(wsOld, wsNew, base, False)
    range = 'A1:J'+str(wsNew.max_row + 10)
    wbNew.create_named_range('_xlnm.Print_Area', wsNew, range, scope=0)

    # save new sheet out to new file
    try:
        wbNew.save(output_name)
    except Exception as e:
        log('             FAILED TO CREATE XLSX: ' + str(e), True)


def process_sheets():
    global clemens
    log("\nPROCESS SHEETS ===============================")
    os.chdir(source_dir + 'downloads/')
    for file in sorted(glob.glob('*.xlsx')):
        is_clemens(file[:7] == 'Clemens')
        log("  converting %s to pdf" % (file))
        process_sheet_to_pdf(file)
        log("  converting %s to xlsx" % (file))
        process_sheet_to_xlsx(file)
        log("")


def download_sheets():
    files = os.listdir(source_dir + 'downloads')
    for file in files:
        os.remove(os.path.join(source_dir + 'downloads', file))

    smart = smartsheet.Smartsheet(api)
    smart.assume_user(os.getenv('SMARTSHEET_USER'))
    log("DOWNLOADING SHEETS ===========================")
    for report in reports:
        log("  downloading sheet: " + report['name'])
        try:
            smart.Reports.get_report_as_excel(report['id'], source_dir + 'downloads')
        except Exception as e:
            log('                     ERROR DOWNLOADING SHEET: ' + str(e), True)

def send_error_report():
    subject = 'Smartsheet Boats on Order Error Report'
    mail_results(subject, log_text)

def main():
    try:
        download_sheets()
        # process_sheets()
    except Exception as e:
        log('Uncaught Error in main(): ' + str(e), True)
    if (errors):
        # send_error_report()
        pass


@click.command()
@click.option(
    '--download/--no-download',
    default=True,
    help='Download sheet(s)'
)
@click.option(
    '--pdf/--no-pdf',
    default=True,
    help='Create PDFs'
)
@click.option(
    '--excel/--no-excel',
    default=True,
    help='Create Excel Sheets'
)
@click.option(
    '--dealer',
    '-d',
    multiple=True,
    help='Dealers to include'
)
@click.option(
    '--ignore',
    '-i',
    multiple=True,
    help='Dealers to ignore'
)
def cli(download, pdf, excel, dealer, ignore):
    """
    stub function here
    """
    print(download)
    # main()

if __name__ == "__main__":
    cli()  # pylint: disable=no-value-for-parameter
