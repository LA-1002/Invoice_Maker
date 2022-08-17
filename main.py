import csv,operator
from fpdf import FPDF;
from datetime import date;
import aspose.words as aw

def GetInvoiceID(num):
    today = date.today();
    today = ('{}-{}'.format(today.month,today.year));
    return ('INVOICE-{}-{}'.format(today,num));

def GetCommissionData():
    comData = csv.reader(open('commission_8-16-2022.csv'),delimiter=',')
    comData = sorted(comData, key=operator.itemgetter(2))
    #GET ALL INFO FROM THE BRANDS AND TURN IT INTO A JSON FILE
    brands = [];
    for d in comData:
        brands.append(d[2]);
    brands = list(dict.fromkeys(brands)) #GET RID OF DUPLICATES IN LIST
    brands.remove('BrandID') #GETS RID OF BRAND TITLE IN LIST FILE
    allInvoices = [];  #LIST OF ALL INVOICES
    for b in brands:
        amount = 0;
        orders = 0;
        id = GetInvoiceID((brands.index(b)) + 1);
        invoice = [];
        for d in comData:
            if (d[2] == b):
                amount+=float(d[3]);
                orders+=1;
                invoice.append({
                    "Amount": round(float(d[3]),2),
                    "Date": d[4],
                    'Time': d[5],
                    'DiscountID': d[1]
                })#ADDING BRAND VERSION TO COMPLETE INVOICE JSON
        allInvoices.append({
            'invoiceID': id,
            "BrandID": b,
            "Amount": round(amount,2),
            "Number": orders,
            "Orders": invoice
        })
    return allInvoices;
    


def editWordDoc(invoice,info):
    #Opening the Template Document
    doc = aw.Document('Templates/Invoice Template.docx')
    #The Find & Replace function of the module I have used
    find = aw.replacing.FindReplaceOptions(aw.replacing.FindReplaceDirection.FORWARD)
    #Business INFORMATION REPLACE
    doc.range.replace('Company Name',info['Name'],find),
    doc.range.replace('Address-street # and name',info['Address']['Street'],find),
    doc.range.replace('Address- town',info['Address']['Town'],find),
    doc.range.replace('Address- county',info['Address']['County'],find),
    doc.range.replace('Address- postcode',info['Address']['Postcode'],find),
    #INVOICE INFORMATION REPLACE
    dates = date.today();
    next = int(dates.month) + 1;
    today = ('{}-{}-{}'.format(dates.day,dates.month,dates.year));
    nMonth = ('{}-{}-{}'.format(dates.day,str(next),dates.year));
    doc.range.replace("Date sent (DD/MM/YY)",today, find);
    doc.range.replace("Date due (DD/MM/YY)",nMonth, find);
    doc.range.replace('£ Total commission ',str(invoice['Amount']),find);
    doc.range.replace('Ref number (INV-YYYY-MM-NNN)',invoice['invoiceID'],find);
    where = str(invoice['invoiceID']) + '.docx';
    print();
    doc.save('Invoices/%s'%where);
    


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




allInvoices = GetCommissionData();
business = {
    'Name': "Alidi's",
    'Address': {
        'Street': 'Fake Street',
        'Town': 'Counterfeit Corner',
        'County': 'Phony District',
        'Postcode': 'SHAM PRE',
    }
    }
print(allInvoices[0]);






    



