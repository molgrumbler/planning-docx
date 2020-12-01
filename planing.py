import docx
import argparse
import datetime
from docx.shared import Pt 
 
def paragraph_replace(paragraph,old,new,style):
    if old in paragraph.text:
        if new[0]=='0':
            old1 = '1'+old
            old2 = '2'+old
            old3 = '3'+old
            if not (old1 in paragraph.text or \
                    old2 in paragraph.text or \
                    old3 in paragraph.text):
                paragraph.text = paragraph.text.replace(old,new)
                paragraph.style = style
                
                
        else:
            paragraph.text = paragraph.text.replace(old,new)
            paragraph.style = style
    return paragraph

def doc_replace(document,old,new,style):
    for tb in document.tables: 
        for column in tb.columns:
            for cell in column.cells:
                for paragraph in cell.paragraphs:
                    paragraph = paragraph_replace(paragraph,old,new,style)
    return document

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("year", type=int,
                    help="make planning for YEAR")
    
    args = parser.parse_args()
    b = datetime.date(2018,12,31)
    numdays = 369
    numdays = 17
    document = docx.Document("templ1.docx")
    document1 = docx.Document("templ2.docx")
    style1 = document1.styles['Normal']
    font1 = style1.font
    font1.name = 'Arial'
    font1.size = Pt(9)

    style = document.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(9)

    month_list_o = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
           'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    month_list = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
           'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
    n = datetime.date(args.year,1,1)
    delta = datetime.timedelta(days=n.weekday())
    nd = n-delta
    for i in range(0,numdays):
        delta = datetime.timedelta(days=i)
        old_day = b+delta
        new_day = nd+delta
        s = str(old_day.day)+' '+month_list_o[old_day.month-1]+' '+str(old_day.year)
        s1 = str(new_day.strftime('%d'))+' '+month_list[new_day.month-1]+' '+str(new_day.year)
        document = doc_replace(document,s,s1,style)
        document1 = doc_replace(document1,s,s1,style1)
        print(s1)
       

    document.save("example.docx")
    document1.save("example1.docx")



