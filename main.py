import csv,operator
from fpdf import FPDF;
from datetime import date;

data = csv.reader(open('results_8-12-2022.csv'),delimiter=',')

data = sorted(data, key=operator.itemgetter(2))





def GetInvoices():
    #GET ALL INFO FROM THE BRANDS AND TURN IT INTO A JSON FILE
    brands = [];
    for d in data:
        brands.append(d[1]);
    
    brands = list(dict.fromkeys(brands)) #GET RID OF DUPLICATES IN LIST
    brands.pop(); #GETS RID OF BRAND TITLE IN LIST FILE

    allInvoices = [];  #LIST OF ALL INVOICES

    for b in brands:
        amount = 0;
        orders = 0;
        invoice = [];
        for d in data:
            if (d[1] == b):
                amount+=float(d[2]);
                orders+=1;
                invoice.append({
                    "Amount": round(float(d[2]),2),
                    "Date": d[3],
                    'Time': d[4]
                })#ADDING BRAND VERSION TO COMPLETE INVOICE JSON
        allInvoices.append({
            "BrandID": b,
            "Amount": round(amount,2),
            "Number": orders,
            "Orders": invoice
        })
        return allInvoices;


def GetInvoiceID(num):
    today = date.today();
    today = ('{}-{}-{}'.format(today.day,today.month,today.year));
    return ('INVOICE:{}-{}'.format(today,num));


def makePDF():
    pdf = FPDF();
    pdf.add_page();
    pdf.set_font('arial','B',13.0)
    number = 0;
    for y in range(28):
        number+=1;
        for x in range(21):
            pdf.set_xy((10),(10))
            pdf.cell(ln=0,h=5.0,align='L', w=0,txt=str('THIS IS'),border=0);
    pdf.output('test3.pdf','F')

id = GetInvoiceID(0);





    



