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
    currentDirectory = os.getcwd()
    os.chdir(currentDirectory)

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
        s1_ic = qt.value * s1_cpp.value * (1 + (1 - float(s1_dv.value)/100) * (stripNonNumeric(s1_lt.value)) + (1 - float(s1_qa.value)/100))

        s2_cpp = sheet['J'+str(i)]
        s2_lt = sheet['L'+str(i)]
        s2_qa = sheet['N'+str(i)]
        s2_dv = sheet['M'+str(i)]
        s2_ic = qt.value * s2_cpp.value * (1 + (1 - float(s2_dv.value)/100) * (stripNonNumeric(s2_lt.value)) + (1 - float(s2_qa.value)/100))
        
        s3_cpp = sheet['P'+str(i)]
        s3_lt = sheet['R'+str(i)]
        s3_qa = sheet['T'+str(i)]
        s3_dv = sheet['S'+str(i)]
        s3_ic = qt.value * s3_cpp.value * (1 + (1 - float(s3_dv.value)/100) * (stripNonNumeric(s3_lt.value)) + (1 - float(s3_qa.value)/100))

        if (s1_ic < s2_ic):
            if (s1_ic < s3_ic):
                bestPrices.append(qt.value*s1_cpp.value)
                exportSheet['A' + str(i)] = partName.value
                exportSheet['B' + str(i)] = qt.value
                exportSheet['C' + str(i)] = 'Supplier 1'
                exportSheet['D' + str(i)] = s1_cpp.value
                exportSheet['E' + str(i)] = s1_lt.value
                exportSheet['F' + str(i)] = s1_qa.value
                exportSheet['G' + str(i)] = s1_dv.value
                exportSheet['H' + str(i)] = qt.value*s1_cpp.value
                exportSheet['I' + str(i)] = s1_ic

            else:
                bestPrices.append(qt.value*s3_cpp.value)
                exportSheet['A' + str(i)] = partName.value
                exportSheet['B' + str(i)] = qt.value
                exportSheet['C' + str(i)] = 'Supplier 3'
                exportSheet['D' + str(i)] = s3_cpp.value
                exportSheet['E' + str(i)] = s3_lt.value
                exportSheet['F' + str(i)] = s3_qa.value
                exportSheet['G' + str(i)] = s3_dv.value
                exportSheet['H' + str(i)] = qt.value*s3_cpp.value
                exportSheet['I' + str(i)] = s3_ic
        else:
            if (s2_ic < s3_ic):
                bestPrices.append(qt.value*s2_cpp.value)
                exportSheet['A' + str(i)] = partName.value
                exportSheet['B' + str(i)] = qt.value
                exportSheet['C' + str(i)] = 'Supplier 2'
                exportSheet['D' + str(i)] = s2_cpp.value
                exportSheet['E' + str(i)] = s2_lt.value
                exportSheet['F' + str(i)] = s2_qa.value
                exportSheet['G' + str(i)] = s2_dv.value
                exportSheet['H' + str(i)] = qt.value*s2_cpp.value
                exportSheet['I' + str(i)] = s2_ic
            else:
                bestPrices.append(qt.value*s3_cpp.value)
                exportSheet['A' + str(i)] = partName.value
                exportSheet['B' + str(i)] = qt.value
                exportSheet['C' + str(i)] = 'Supplier 3'
                exportSheet['D' + str(i)] = s3_cpp.value
                exportSheet['E' + str(i)] = s3_lt.value
                exportSheet['F' + str(i)] = s3_qa.value
                exportSheet['G' + str(i)] = s3_dv.value
                exportSheet['H' + str(i)] = qt.value*s3_cpp.value
                exportSheet['I' + str(i)] = s3_ic

    exportSheet['F20'] = '=AVERAGE(F2:F19)'      
    exportSheet['G20'] = '=AVERAGE(G2:G19)'
    exportSheet['H20'] = '=SUM(H2:H19)'
    exportSheet['I20'] = '=SUM(I2:I19)'
    
    for cellObjects in exportSheet['H2:I20']:
        for cell in cellObjects:
            cell.number_format = '$#,###.00'
    
    exportData.save(currentDirectory + '/outputs/test.xlsx')

    for i in range(len(bestPrices)):
        totSum = totSum + bestPrices[i]
    print("Total Price: " + as_currency(totSum))            

if __name__ == "__main__":
    main()



























































