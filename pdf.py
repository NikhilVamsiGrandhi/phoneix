import fitz

def get_value(list,keyword):
    str = list[list.index(keyword)+2]
    str = str.replace('-','').replace(',','')
    float_value = float(str)
    value = int(round(float_value))
    return value

def value_from_pdf(keyword,file_path,total=False):
    doc = fitz.open(rf"{file_path}")
    for page_num in range(doc.page_count):
        page = doc[page_num]
        text = page.get_text()
        list = text.split('\n')
        if total:
            if "TOTAL" in list:
                return get_value(list,keyword)
        else:
            if keyword in list:
                return get_value(list,keyword)
    doc.close()
    return None

# print(value_from_pdf("TRANSPORT",r"C:\Users\NikhilVamsiGrandhi\pdfslocal\MTD_Product_Purchase_By_Date[11_Mar].pdf"))
# print(value_from_pdf("TOTAL",r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PSG\downlaods\MTD_Product_Purchase_By_Date[27_Feb].pdf"))
# print(value_from_pdf("N",r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PSG\downlaods\MTD_Product_Purchase_By_Date[27_Feb].pdf"))
# print(value_from_pdf("TRANSPORT",r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PSG\downlaods\MTD_Product_Purchase_By_Date[27_Feb].pdf"))
# print(value_from_pdf(r"N",r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PSG\downlaods\MTD_Product_Purchase_By_Date[27_Feb].pdf",True))
# print(value_from_pdf(r"B",r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PSG\downlaods\MTD_Product_Purchase_By_Date[27_Feb].pdf",True))