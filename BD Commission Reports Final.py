import pandas
import locale
import datetime

minDate = datetime.date(2020, 5, 24)
maxDate = datetime.date(2020, 5, 30)

locale.setlocale( locale.LC_ALL, 'en_US.UTF-8' ) 

def Date(temp):
    month = int(temp[0:2])
    day = int(temp[3:5])
    year = int(temp[6:10])
    a = datetime.date(year, month, day)
    return a

def commAmount(df):
    total = 0
    for i in range(len(df)):
        saleAmount = df.loc[i, 'Amount']
        total += saleAmount
    return total

def commRate(df):
    arr = []
    c = 0
    for i in range(len(df)):
        saleType = df.loc[i, 'Type']
        bid = df.loc[i, 'BID']
        if saleType == 'Rejected/Invalid Payment':
            #c += 1
            pass
        else:
            arr.append(bid)
    arr2 = []
    for i in arr:
        if i in arr2:
            pass
        else:
            arr2.append(i)
    count = len(arr2)
    if count <= 2:
        rate = .35
    elif count <= 4:
        rate = .4
    elif count <= 6:
        rate = .45
    elif count <= 8:
        rate = .5
    elif count == 9:
        rate = .53
    elif count >= 10:
        rate = .55
    else:
        print('Illegal argument')
    arr = []
    #print(count, rate)
    return rate

def manCommRate(df):
    return .5

def seniorCommRate(df):
    count = 0
    for i in range(len(df)):
        saleType = df.loc[i, 'Type']
        if saleType == 'Rejected/Invalid Payment':
            count -= 1
        else:
            count += 1
    if count <= 2:
        rate = .35
    elif count <= 4:
        rate = .4
    elif count <= 6:
        rate = .5
    elif count <= 8:
        rate = .55
    elif count == 9:
        rate = .6
    elif count >= 10:
        rate = .65
    else:
        print('Illegal argument')
    return rate
    
def totalComm(amount, rate):
    commission = amount * rate
    return commission
    

data = pandas.read_csv('C:/Users/jkreuger/Downloads/PaymentReport (18).csv', error_bad_lines=False)
columnNames = ['BID', 'Invoice', 'Business', 'Joined', 'Billed', 'Item', 'Payment', 'Item Amt', 'Payment Date', 'Type', 'Notes', 'Sales', 'Retention Rep' ]
genPayReport = pandas.DataFrame(columns = columnNames)

testColumns = ['BID', 'Invoice']
testReport = pandas.DataFrame(columns = columnNames)

i = 0
j = 1
k = -1


for row in data.iterrows():
    sales = data.loc[i, 'Sales']
    if sales == ' ':
        sales = sales.replace(' ', 'House')
    bid = data.loc[i, 'BID']
    invoice = data.loc[i, 'Invoice']
    business = data.loc[i, 'Business']
    joined = data.loc[i, 'Joined']
    billed = data.loc[i, 'Billed']
    item = data.loc[i, 'Item']
    payment = data.loc[i, 'Payment']
    itemAmt = data.loc[i, 'Item Amt']
    payDate = data.loc[i, 'Payment Date']
    type0 = data.loc[i, 'Type']
    notes = data.loc[i, 'Notes']
    #sales = data.loc[i, 'Sales']
    retRep = data.loc[i, 'Retention Rep                        ']
    #retRep = 0
    testReport = testReport.append({'BID': bid, 'Invoice': invoice, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment': payment, 'Item Amt': itemAmt, 'Payment Date': payDate, 'Type': type0, 'Notes': notes, 'Sales': sales, 'Retention Rep': retRep}, ignore_index = True)
    #genPayReport = genPayReport.append({'BID': bid, 'Invoice': invoice, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment': payment, 'Item Amt': itemAmt, 'Payment Date': payDate, 'Type': type0, 'Notes': notes, 'Sales': sales}, ingore_index = True)
    i += 1
    j += 1
    k += 1
    
testReport.to_excel("C:/Users/jkreuger/Downloads/GenPayReport.xlsx")


#testReport['Sales'].fillna('House') 
#testReport = testReport.sort_values(by=['Sales'])

sumColumns = ['BID', 'Business', 'Joined', 'Billed', 'Item', 'Payment Date', 'Type', 'Notes', 'Sales', 'Amount']
summaryGonzalez = pandas.DataFrame(columns = sumColumns)
summaryMckethan = pandas.DataFrame(columns = sumColumns)
summaryWest = pandas.DataFrame(columns = sumColumns)
summaryDeLaCruz = pandas.DataFrame(columns = sumColumns)
summaryBrown = pandas.DataFrame(columns = sumColumns)
summaryPellak = pandas.DataFrame(columns = sumColumns)
summaryMcCance = pandas.DataFrame(columns = sumColumns)
summaryMcAdams = pandas.DataFrame(columns = sumColumns)
summaryGSmith = pandas.DataFrame(columns = sumColumns)
summaryDudo = pandas.DataFrame(columns = sumColumns)
summaryVonVogt = pandas.DataFrame(columns = sumColumns)
summaryDawson = pandas.DataFrame(columns = sumColumns)
summaryAtkins = pandas.DataFrame(columns = sumColumns)
summaryJoffrion = pandas.DataFrame(columns = sumColumns)
summaryRifenburg = pandas.DataFrame(columns = sumColumns)
summaryAcosta = pandas.DataFrame(columns = sumColumns)
summaryFox = pandas.DataFrame(columns = sumColumns)
summaryYokom = pandas.DataFrame(columns = sumColumns)
summaryBononcini = pandas.DataFrame(columns = sumColumns)
summaryRobberson = pandas.DataFrame(columns = sumColumns)
summaryBacon = pandas.DataFrame(columns = sumColumns)
summaryChevere = pandas.DataFrame(columns = sumColumns)
summarySSmith = pandas.DataFrame(columns = sumColumns)
summaryLewis = pandas.DataFrame(columns = sumColumns)
summaryColeman = pandas.DataFrame(columns = sumColumns)
summaryOverton = pandas.DataFrame(columns = sumColumns)
summaryWagoner = pandas.DataFrame(columns = sumColumns)
summaryFerrigno = pandas.DataFrame(columns = sumColumns)
summaryBelford = pandas.DataFrame(columns = sumColumns)
summaryBarlow = pandas.DataFrame(columns = sumColumns)
summaryYsasi = pandas.DataFrame(columns = sumColumns)


i = 0
j = 1

for i in range(len(testReport)):
    bid = testReport.loc[i, 'BID']
    business = testReport.loc[i, 'Business']
    joined = testReport.loc[i, 'Joined']
    if joined == ' ':
        continue
    else:
        joinDate = Date(joined)
        if joinDate <= minDate or joinDate >= maxDate:
            continue
        elif testReport.loc[i, 'Type'] == 'Write Off':
            continue
    billed = testReport.loc[i, 'Billed']
    item = testReport.loc[i, 'Item']
    payDate = testReport.loc[i, 'Payment Date']
    type0 = testReport.loc[i, 'Type']
    notes = testReport.loc[i, 'Notes']
    salesRep = testReport.loc[i, 'Sales']
    rawAmt = testReport.loc[i, 'Item Amt']
    amt1 = rawAmt[1:]
    amt = locale.atof(amt1)
    if item == 'Dues (New)' or item == 'Additional Business(es)' or item == 'Logo Package' or item == 'Additional Location(s)':
        pass
    else:
        continue
    if type0 == 'Rejected/Invalid Payment':
        amt = 0 - amt
    if salesRep == 'Brandi Gonzalez':
        summaryGonzalez = summaryGonzalez.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Dale Mckethan' or salesRep == 'Dale McKethan':
        summaryMckethan = summaryMckethan.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Damon West':
        summaryWest = summaryWest.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Danielle De La Cruz':
        summaryDeLaCruz = summaryDeLaCruz.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Donna Brown':
        summaryBrown = summaryBrown.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Donna Pellak':
        summaryPellak = summaryPellak.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Elaine McCance':
        summaryMcCance = summaryMcCance.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Galadriel McAdams':
        summaryMcAdams = summaryMcAdams.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Gary Smith':
        summaryGSmith = summaryGSmith.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Jack Dudo':
        summaryDudo = summaryDudo.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Janice Von Vogt':
        summaryVonVogt = summaryVonVogt.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Jessica Dawson':
        summaryDawson = summaryDawson.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Jonathon Atkins' or salesRep == 'Jonathan Atkins':
        summaryAtkins = summaryAtkins.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Kelly Overton':
        summaryOverton = summaryOverton.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Katie Joffrion':
        summaryJoffrion = summaryJoffrion.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Kim Rifenburg':
        summaryRifenburg = summaryRifenburg.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Mario Acosta':
        summaryAcosta = summaryAcosta.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Meghan Fox':
        summaryFox = summaryFox.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Michael Yokom' or salesRep == 'Mike Yokom':
        summaryYokom = summaryYokom.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Michelle Bononcini':
        summaryBononcini = summaryBononcini.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Monica Robberson':
        summaryRobberson = summaryRobberson.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Paul Bacon':
        summaryBacon = summaryBacon.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)    
    elif salesRep == 'Richard Chevere':
        summaryChevere = summaryChevere.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Rose Wagoner':
        summaryWagoner = summaryWagoner.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Shawn Ferrigno':
        summaryFerrigno = summaryFerrigno.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Sheila Belford':
        summaryBelford = summaryBelford.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Stephen Smith':
        summarySSmith = summarySSmith.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)  
    elif salesRep == 'Tony Barlow':
        summaryBarlow = summaryBarlow.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Von Ysasi':
        summaryYsasi = summaryYsasi.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Tori Lewis':
        summaryLewis = summaryLewis.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Wanda  Coleman':
        summaryColeman = summaryColeman.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)


#print(totalComm(summaryGSmith))
#print(summaryGSmith.head(10))

commissionColumns = ['Rep', 'Office', 'Commissionable Amount', 'Commission Rate', 'Earned Commission', 'COVID-19 Minimum']
commissionBD = pandas.DataFrame(columns = commissionColumns)

#commissionBD = commissionBD.append({'Rep': 'Dale Mckethan', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryMckethan), 'Commission Rate': commRate(summaryMckethan), 'Earned Commission': totalComm(commRate(summaryMckethan), commAmount(summaryMckethan))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Damon West', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryWest), 'Commission Rate': seniorCommRate(summaryWest), 'Earned Commission': totalComm(seniorCommRate(summaryWest), commAmount(summaryWest))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Donna Brown', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryBrown), 'Commission Rate': commRate(summaryBrown), 'Earned Commission': totalComm(commRate(summaryBrown), commAmount(summaryBrown))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Donna Pellak', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryPellak), 'Commission Rate': commRate(summaryPellak), 'Earned Commission': totalComm(commRate(summaryPellak), commAmount(summaryPellak))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Elaine McCance', 'Office': 'San Antonio', 'Commissionable Amount': commAmount(summaryMcCance), 'Commission Rate': commRate(summaryMcCance), 'Earned Commission': totalComm(commRate(summaryMcCance), commAmount(summaryMcCance))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Gary Smith', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryGSmith), 'Commission Rate': commRate(summaryGSmith), 'Earned Commission': totalComm(commRate(summaryGSmith), commAmount(summaryGSmith))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Jack Dudo', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryDudo), 'Commission Rate': commRate(summaryDudo), 'Earned Commission': totalComm(commRate(summaryDudo), commAmount(summaryDudo))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Janice Von Vogt', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryVonVogt), 'Commission Rate': commRate(summaryVonVogt), 'Earned Commission': totalComm(commRate(summaryVonVogt), commAmount(summaryVonVogt))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Jessica Dawson', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryDawson), 'Commission Rate': manCommRate(summaryDawson), 'Earned Commission': totalComm(manCommRate(summaryDawson), commAmount(summaryDawson))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Jonathon Atkins', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryAtkins), 'Commission Rate': commRate(summaryAtkins), 'Earned Commission': totalComm(commRate(summaryAtkins), commAmount(summaryAtkins))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Kelly Overton', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryOverton), 'Commission Rate': commRate(summaryOverton), 'Earned Commission': totalComm(commRate(summaryOverton), commAmount(summaryOverton))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Kim Rifenburg', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryRifenburg), 'Commission Rate': commRate(summaryRifenburg), 'Earned Commission': totalComm(commRate(summaryRifenburg), commAmount(summaryRifenburg))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Mario Acosta', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryAcosta), 'Commission Rate': commRate(summaryAcosta), 'Earned Commission': totalComm(commRate(summaryAcosta), commAmount(summaryAcosta))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Michelle Bononcini', 'Office': 'San Antonio', 'Commissionable Amount': commAmount(summaryBononcini), 'Commission Rate': commRate(summaryBononcini), 'Earned Commission': totalComm(commRate(summaryBononcini), commAmount(summaryBononcini))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Michael Yokom', 'Office': 'San Antonio', 'Commissionable Amount': commAmount(summaryYokom), 'Commission Rate': commRate(summaryYokom), 'Earned Commission': totalComm(commRate(summaryYokom), commAmount(summaryYokom))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Richard Chevere', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryChevere), 'Commission Rate': commRate(summaryChevere), 'Earned Commission': totalComm(commRate(summaryChevere), commAmount(summaryChevere))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Rose Wagoner', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryWagoner), 'Commission Rate': commRate(summaryWagoner), 'Earned Commission': totalComm(commRate(summaryWagoner), commAmount(summaryWagoner))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Shawn Ferrigno', 'Office': 'San Antonio', 'Commissionable Amount': commAmount(summaryFerrigno), 'Commission Rate': commRate(summaryFerrigno), 'Earned Commission': totalComm(commRate(summaryFerrigno), commAmount(summaryFerrigno))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Sheila Belford', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryBelford), 'Commission Rate': commRate(summaryBelford), 'Earned Commission': totalComm(commRate(summaryBelford), commAmount(summaryBelford))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Stephern Smith', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summarySSmith), 'Commission Rate': commRate(summarySSmith), 'Earned Commission': totalComm(commRate(summarySSmith), commAmount(summarySSmith))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Tony Barlow', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryBarlow), 'Commission Rate': manCommRate(summaryBarlow), 'Earned Commission': totalComm(manCommRate(summaryBarlow), commAmount(summaryBarlow))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Von Ysasi', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryYsasi), 'Commission Rate': commRate(summaryYsasi), 'Earned Commission': totalComm(commRate(summaryYsasi), commAmount(summaryYsasi))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Wanda Coleman', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryColeman), 'Commission Rate': commRate(summaryColeman), 'Earned Commission': totalComm(commRate(summaryColeman), commAmount(summaryColeman))}, ignore_index=True)

newSalesList = [summaryMckethan, summaryWest, summaryBrown, summaryPellak, summaryMcCance, summaryGSmith, summaryDudo, summaryVonVogt, summaryDawson, summaryAtkins, summaryOverton, summaryRifenburg, summaryAcosta, summaryBononcini, summaryYokom, summaryChevere, summaryWagoner, summaryFerrigno, summaryBelford, summarySSmith, summaryBarlow, summaryYsasi, summaryColeman]

    
austinList = [summaryBrown, summaryPellak, summaryGSmith, summaryDudo, summaryDawson, summaryRifenburg, summaryAcosta, summaryYsasi, summaryColeman]
fortWorthList = [summaryMckethan, summaryWest, summaryVonVogt, summaryAtkins, summaryOverton, summaryBelford, summarySSmith, summaryBarlow]
sanAntonioList = [summaryMcCance, summaryBononcini, summaryYokom, summaryFerrigno]

masterList = [austinList, fortWorthList, sanAntonioList]

writer = pandas.ExcelWriter('C:/Users/jkreuger/Downloads/report test.xlsx', engine='xlsxwriter')

reportList = []

for i in masterList:
    for j in i:
        print(j)
        print(commAmount(j), commRate(j), totalComm(commRate(j), commAmount(j)))

        reportList.append(j)

managerColumns = ['Item', 'Type', 'Notes', 'Processed', 'Sales', 'Item Amt']
austinManager = pandas.DataFrame(columns = managerColumns)
fortWorthManager = pandas.DataFrame(columns = managerColumns)
sanAntonioManager = pandas.DataFrame(columns = managerColumns)

df = summaryMckethan.append(summaryWest)
df = df.append(summaryBrown)
df = df.append(summaryPellak)
df = df.append(summaryMcCance)
df = df.append(summaryGSmith)
df = df.append(summaryDudo)
df = df.append(summaryVonVogt)
df = df.append(summaryDawson)
df = df.append(summaryAtkins)
df = df.append(summaryOverton)
df = df.append(summaryRifenburg)
df = df.append(summaryAcosta)
df = df.append(summaryBononcini)
df = df.append(summaryYokom)
df = df.append(summaryChevere)
df = df.append(summaryWagoner)
df = df.append(summaryFerrigno)
df = df.append(summaryBelford)
df = df.append(summarySSmith)
df = df.append(summaryBarlow)
df = df.append(summaryYsasi)
df = df.append(summaryColeman)


commissionBD.to_excel(writer, sheet_name = 'Summary', index = False)
df.to_excel(writer, sheet_name='New', index = False)
writer.save()




#summaryBelford.to_excel(writer, sheet_name = 'Belford')
#print(commissionBD)

#print(summaryDudo)

#print(commAmount(summaryDudo), commRate(summaryDudo), totalComm(commRate(summaryDudo), commAmount(summaryDudo)))


"""
writer = pandas.ExcelWriter('C:/Users/jkreuger/Downloads/wb test.xlsx', engine='xlsxwriter')

testReport.to_excel(writer, sheet_name='Sheet1')
commissionBD.to_excel(writer, sheet_name='Sheet2')
summarySSmith.to_excel(writer, sheet_name='Sheet3')
"""

