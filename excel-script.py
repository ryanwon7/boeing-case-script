import openpyxl, os, re

def stripNonNumeric(inputString):
    output = re.sub("[a-z]", "", inputString, flags=re.IGNORECASE)
    outputNum = float(output)
    return outputNum

def as_currency(amount):
    if amount >= 0:
        return '${:,.2f}'.format(amount)
    else:
        return '-${:,.2f}'.format(-amount)

def main():
    os.chdir('C:/Users/ryanwon7/Desktop/boeing-case-script/')
    importData = openpyxl.load_workbook('suppliers.xlsx')
    exportData = openpyxl.Workbook()

    sheet = importData.get_sheet_by_name('Major Suppliers')
    exportSheet = exportData.active
    
    exportSheet.title = 'Exported Data'
    exportSheet['A1'] = 'Part'
    exportSheet['B1'] = 'Quantity'
    exportSheet['C1'] = 'Supplier Number'
    exportSheet['D1'] = 'Price per Part'
    exportSheet['E1'] = 'Lead Time'
    exportSheet['F1'] = 'Quality Acceptance'
    exportSheet['G1'] = 'On-Time Delivery'
    exportSheet['H1'] = 'Total Price'
    exportSheet['I1'] = 'Impact Price'

    totSum = 0 #totalPrice
    bestPrices = []
    partName = 0
    rowLetter = ''
    qt = 0 #number of items per plane for a part
    s1_cpp = 0 #Supplier 1 - Cost per part (dollars)
    s1_lt = 0 #Supplier 1 - lead time
    s1_qa = 0 #Supplier 1 - quality acceptance (decimal percentage)
    s1_dv = 0 #Supplier 1 - delivery percentage
    s2_cpp = 0 #Supplier 2 - Cost per part (dollars)
    s2_lt = 0 #Supplier 2 - lead time
    s2_qa = 0 #Supplier 2 - quality acceptance (decimal percentage)
    s2_dv = 0 #Supplier 2 - delivery percentage
    s3_cpp = 0 #Supplier 3 - Cost per part (dollars)
    s3_lt = 0 #Supplier 3 - lead time
    s3_qa = 0 #Supplier 3 - quality acceptance (decimal percentage)
    s3_dv = 0 #Supplier 3 - delivery percentage

    s1_ic = 0 #Model calculated "Impact Cost" of the part from Supplier 1
    s2_ic = 0 #Model calculated "Impact Cost" of the part from Supplier 2
    s3_ic = 0 #Model calculated "Impact Cost" of the part from Supplier 3

    for i in range(2, 20):
        partName = sheet['A'+str(i)]
        qt = sheet['B'+str(i)]

        s1_cpp = sheet['D'+str(i)]
        s1_lt = sheet['F'+str(i)]
        s1_qa = sheet['H'+str(i)]
        s1_dv = sheet['G'+str(i)]
        s1_ic = qt.value * s1_cpp.value * (1 + (1 - float(s1_dv.value)/100) * (stripNonNumeric(s1_lt.value))/4 + (1 - float(s1_qa.value)/100))

        s2_cpp = sheet['J'+str(i)]
        s2_lt = sheet['L'+str(i)]
        s2_qa = sheet['N'+str(i)]
        s2_dv = sheet['M'+str(i)]
        s2_ic = qt.value * s2_cpp.value * (1 + (1 - float(s2_dv.value)/100) * (stripNonNumeric(s2_lt.value))/4 + (1 - float(s2_qa.value)/100))
        
        s3_cpp = sheet['P'+str(i)]
        s3_lt = sheet['R'+str(i)]
        s3_qa = sheet['T'+str(i)]
        s3_dv = sheet['S'+str(i)]
        s3_ic = qt.value * s3_cpp.value * (1 + (1 - float(s3_dv.value)/100) * (stripNonNumeric(s3_lt.value))/4 + (1 - float(s3_qa.value)/100))

        if (s1_ic < s2_ic):
            if (s1_ic < s3_ic):
                bestPrices.append(qt.value*s1_cpp.value)
                exportSheet['A' + '1'] = partName.value
                exportSheet['A' + '2'] = qt.value
                exportSheet[rowLetter + '3'] = 'Supplier 1'
                exportSheet[rowLetter + '4'] = s1_cpp.value
                exportSheet[rowLetter + '5'] = s1_lt.value
                exportSheet[rowLetter + '6'] = s1_qa.value
                exportSheet[rowLetter + '7'] = s1_dv.value
                exportSheet[rowLetter + '8'] = qt.value*s1_cpp.value
                exportSheet[rowLetter + '9'] = s1_ic

            else:
                bestPrices.append(qt.value*s3_cpp.value)
                exportSheet[rowLetter + '1'] = partName.value
                exportSheet[rowLetter + '2'] = qt.value
                exportSheet[rowLetter + '3'] = 'Supplier 3'
                exportSheet[rowLetter + '4'] = s3_cpp.value
                exportSheet[rowLetter + '5'] = s3_lt.value
                exportSheet[rowLetter + '6'] = s3_qa.value
                exportSheet[rowLetter + '7'] = s3_dv.value
                exportSheet[rowLetter + '8'] = qt.value*s3_cpp.value
                exportSheet[rowLetter + '9'] = s3_ic
        else:
            if (s2_ic < s3_ic):
                bestPrices.append(qt.value*s2_cpp.value)
                exportSheet[rowLetter + '1'] = partName.value
                exportSheet[rowLetter + '2'] = qt.value
                exportSheet[rowLetter + '3'] = 'Supplier 2'
                exportSheet[rowLetter + '4'] = s2_cpp.value
                exportSheet[rowLetter + '5'] = s2_lt.value
                exportSheet[rowLetter + '6'] = s2_qa.value
                exportSheet[rowLetter + '7'] = s2_dv.value
                exportSheet[rowLetter + '8'] = qt.value*s2_cpp.value
                exportSheet[rowLetter + '9'] = s2_ic
            else:
                bestPrices.append(qt.value*s3_cpp.value)
                exportSheet[rowLetter + '1'] = partName.value
                exportSheet[rowLetter + '2'] = qt.value
                exportSheet[rowLetter + '3'] = 'Supplier 3'
                exportSheet[rowLetter + '4'] = s3_cpp.value
                exportSheet[rowLetter + '5'] = s3_lt.value
                exportSheet[rowLetter + '6'] = s3_qa.value
                exportSheet[rowLetter + '7'] = s3_dv.value
                exportSheet[rowLetter + '8'] = qt.value*s3_cpp.value
                exportSheet[rowLetter + '9'] = s3_ic
                

    for i in range(len(bestPrices)):
        totSum = totSum + bestPrices[i]
    print(as_currency(totSum))

    exportData.save('scriptoutput2.xlsx')
                

if __name__ == "__main__":
    main()












































