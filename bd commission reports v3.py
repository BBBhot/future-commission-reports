import pandas as pd
import locale
import datetime
import os
import shutil
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

today = datetime.date.today() 
dayOfWeek = datetime.datetime.today().weekday()
low = dayOfWeek + 8
up = dayOfWeek + 1
minDate = datetime.date.today() - datetime.timedelta(days = low)
maxDate = datetime.date.today() - datetime.timedelta(days = up)

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
    for i in range(len(df)):
        saleType = df.loc[i, 'Type']
        bid = df.loc[i, 'BID']
        if saleType != 'Rejected/Invalid Payment' and bid not in arr:
            arr.append(bid)
    count = len(arr)
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
    return rate

def manCommRate(df):
    return .5

def seniorCommRate(df):
    arr = []
    for i in range(len(df)):
        saleType = df.loc[i, 'Type']
        bid = df.loc[i, 'BID']
        if saleType != 'Rejected/Invalid Payment' and bid not in arr:
            arr.append(bid)
    count = len(arr)
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

def managerTotal(office):
    totalCount = 0
    commCount = 0
    if office == 'Austin':
        for i, row in austinDf.iterrows():
            tot = austinDf.loc[i, 'Total']
            com = austinDf.loc[i, 'Commission']
            rep = austinDf.loc[i, 'Sales']
            if pd.isna(tot) or rep == 'Jessica Dawson':
                continue
            print(tot, com)
            totalCount += tot
            commCount += com
        print('Jessica', totalCount, totalCount*.03, commCount)
    if office == 'FwSa':
        for i, row in fw_sa_df.iterrows():
            tot = fw_sa_df.loc[i, 'Total']
            com = fw_sa_df.loc[i, 'Commission']
            rep = fw_sa_df.loc[i, 'Sales']
            if pd.isna(tot) or rep == 'Tony Barlow':
                continue
            print(tot, com)
            totalCount += tot
            commCount += com
        print('TOny', totalCount, totalCount*.045, commCount)
    return (totalCount, commCount)

def ReformatDate(date):
    date = str(date)
    temp = date.split('-')
    newDate = temp[1] + '/' + temp[2] + '/' + temp[0]
    return newDate

driver = webdriver.Chrome(r'C:\Users\jkreuger\Downloads\chromedriver_win32\chromedriver.exe')
url = 'https://blue.bbb.org/Core/Login.aspx?ReturnUrl=%2fcore%2fcollections%2fpayments.aspx'
driver.get(url)

driver.implicitly_wait(10)

username = driver.find_element_by_id('l_UserName')
password = driver.find_element_by_id('l_Password')

username.send_keys('jkrueger@austin.bbb.org')
password.send_keys('BBBhot2020!')

driver.find_element_by_id('l_LoginButton').click()

driver.implicitly_wait(10)

begin = driver.find_element_by_id('ctl00_c1_dteP1_dateInput_dateInput')
end = driver.find_element_by_id('ctl00_c1_dteP2_dateInput_dateInput')

begin.clear()
end.clear()

begin.send_keys(ReformatDate(minDate))
end.send_keys(ReformatDate(maxDate))
driver.implicitly_wait(5)
end.send_keys(Keys.ENTER)

driver.implicitly_wait(10)

driver.find_element_by_id('ctl00_c1_ex2_iCSV').click()

time.sleep(10)

filename = max(['C:/Users/jkreuger/Downloads' + "\\" + f for f in os.listdir('C:/Users/jkreuger/Downloads')],key=os.path.getctime)
shutil.move(filename,os.path.join('C:/Users/jkreuger/Downloads',r"payreport{}.csv".format(today)))


#driver.close()

data = pd.read_csv('C:/Users/jkreuger/Downloads/payreport{}.csv'.format(today), error_bad_lines=False)
data.drop(columns = ['Type'])
columnNames = ['BID', 'Invoice', 'Business', 'Joined', 'Billed', 'Item', 'Payment', 'Item Amt', 'Payment Date', 'Type', 'Notes', 'Sales', 'Retention Rep' ]
genPayReport = pd.DataFrame(columns = columnNames)
testColumns = ['BID', 'Invoice']
testReport = pd.DataFrame(columns = columnNames)

for i, row in data.iterrows():
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
    retRep = data.loc[i, 'Retention Rep                        ']
    testReport = testReport.append({'BID': bid, 'Invoice': invoice, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment': payment, 'Item Amt': itemAmt, 'Payment Date': payDate, 'Type': type0, 'Notes': notes, 'Sales': sales, 'Retention Rep': retRep}, ignore_index = True)

sumColumns = ['BID', 'Business', 'Joined', 'Billed', 'Item', 'Payment Date', 'Type', 'Notes', 'Sales', 'Amount']
summaryGonzalez = pd.DataFrame(columns = sumColumns)
summaryMckethan = pd.DataFrame(columns = sumColumns)
summaryWest = pd.DataFrame(columns = sumColumns)
summaryDeLaCruz = pd.DataFrame(columns = sumColumns)
summaryBrown = pd.DataFrame(columns = sumColumns)
summaryPellak = pd.DataFrame(columns = sumColumns)
summaryMcCance = pd.DataFrame(columns = sumColumns)
summaryMcAdams = pd.DataFrame(columns = sumColumns)
summaryGSmith = pd.DataFrame(columns = sumColumns)
summaryDudo = pd.DataFrame(columns = sumColumns)
summaryVonVogt = pd.DataFrame(columns = sumColumns)
summaryDawson = pd.DataFrame(columns = sumColumns)
summaryAtkins = pd.DataFrame(columns = sumColumns)
summaryJoffrion = pd.DataFrame(columns = sumColumns)
summaryRifenburg = pd.DataFrame(columns = sumColumns)
summaryAudagnotti = pd.DataFrame(columns = sumColumns)
summaryAcosta = pd.DataFrame(columns = sumColumns)
summaryFox = pd.DataFrame(columns = sumColumns)
summaryLaws = pd.DataFrame(columns = sumColumns)
summaryYokom = pd.DataFrame(columns = sumColumns)
summaryBononcini = pd.DataFrame(columns = sumColumns)
summaryRobberson = pd.DataFrame(columns = sumColumns)
summaryBacon = pd.DataFrame(columns = sumColumns)
summaryChevere = pd.DataFrame(columns = sumColumns)
summarySSmith = pd.DataFrame(columns = sumColumns)
summaryLewis = pd.DataFrame(columns = sumColumns)
summaryColeman = pd.DataFrame(columns = sumColumns)
summaryOverton = pd.DataFrame(columns = sumColumns)
summaryWagoner = pd.DataFrame(columns = sumColumns)
summaryFerrigno = pd.DataFrame(columns = sumColumns)
summaryBelford = pd.DataFrame(columns = sumColumns)
summaryBarlow = pd.DataFrame(columns = sumColumns)
summaryYsasi = pd.DataFrame(columns = sumColumns)
summaryCanchola = pd.DataFrame(columns = sumColumns)
summaryDore = pd.DataFrame(columns = sumColumns)
summaryHilburn = pd.DataFrame(columns = sumColumns)
summaryBonds = pd.DataFrame(columns = sumColumns)



for i, row in testReport.iterrows():
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
    elif salesRep == 'Carolyn Canchola':
        summaryCanchola = summaryCanchola.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
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
    elif salesRep == 'Lydon Audagnotti':
        summaryAudagnotti = summaryAudagnotti.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Mario Acosta':
        summaryAcosta = summaryAcosta.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Meghan Fox':
        summaryFox = summaryFox.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Melissa Laws':
        summaryLaws = summaryLaws.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Mike Dore' or salesRep == 'Michael Dore':
        summaryDore = summaryDore.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
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
    elif salesRep == 'Robert Bonds':
        summaryBonds = summaryBonds.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Rose Wagoner':
        summaryWagoner = summaryWagoner.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Shawn Ferrigno':
        summaryFerrigno = summaryFerrigno.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
    elif salesRep == 'Shayne Hilburn':
        summaryHilburn = summaryHilburn.append({'BID': bid, 'Business': business, 'Joined': joined, 'Billed': billed, 'Item': item, 'Payment Date': payDate, 'Type': type0, 'Notes': notes,'Sales': salesRep, 'Amount': amt}, ignore_index=True)
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

commissionColumns = ['Rep', 'Office', 'Commissionable Amount', 'Commission Rate', 'Earned Commission', 'COVID-19 Minimum']
commissionBD = pd.DataFrame(columns = commissionColumns)

commissionBD = commissionBD.append({'Rep': 'Carolyn Canchola', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryCanchola), 'Commission Rate': commRate(summaryCanchola), 'Earned Commission': totalComm(commRate(summaryCanchola), commAmount(summaryCanchola))}, ignore_index=True)
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
commissionBD = commissionBD.append({'Rep': 'Lydon Audagnotti', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryAudagnotti), 'Commission Rate': commRate(summaryAudagnotti), 'Earned Commission': totalComm(commRate(summaryAudagnotti), commAmount(summaryAudagnotti))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Mario Acosta', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryAcosta), 'Commission Rate': commRate(summaryAcosta), 'Earned Commission': totalComm(commRate(summaryAcosta), commAmount(summaryAcosta))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Melissa Laws', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryLaws), 'Commission Rate': commRate(summaryLaws), 'Earned Commission': totalComm(commRate(summaryLaws), commAmount(summaryLaws))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Michelle Bononcini', 'Office': 'San Antonio', 'Commissionable Amount': commAmount(summaryBononcini), 'Commission Rate': commRate(summaryBononcini), 'Earned Commission': totalComm(commRate(summaryBononcini), commAmount(summaryBononcini))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Michael Yokom', 'Office': 'San Antonio', 'Commissionable Amount': commAmount(summaryYokom), 'Commission Rate': commRate(summaryYokom), 'Earned Commission': totalComm(commRate(summaryYokom), commAmount(summaryYokom))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Mike Dore', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryDore), 'Commission Rate': commRate(summaryDore), 'Earned Commission': totalComm(commRate(summaryDore), commAmount(summaryDore))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Richard Chevere', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryChevere), 'Commission Rate': commRate(summaryChevere), 'Earned Commission': totalComm(commRate(summaryChevere), commAmount(summaryChevere))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Robert Bonds', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryBonds), 'Commission Rate': commRate(summaryBonds), 'Earned Commission': totalComm(commRate(summaryBonds), commAmount(summaryBonds))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Rose Wagoner', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryWagoner), 'Commission Rate': commRate(summaryWagoner), 'Earned Commission': totalComm(commRate(summaryWagoner), commAmount(summaryWagoner))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Shawn Ferrigno', 'Office': 'San Antonio', 'Commissionable Amount': commAmount(summaryFerrigno), 'Commission Rate': commRate(summaryFerrigno), 'Earned Commission': totalComm(commRate(summaryFerrigno), commAmount(summaryFerrigno))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Shayne Hilburn', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryHilburn), 'Commission Rate': commRate(summaryHilburn), 'Earned Commission': totalComm(commRate(summaryHilburn), commAmount(summaryHilburn))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Sheila Belford', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryBelford), 'Commission Rate': commRate(summaryBelford), 'Earned Commission': totalComm(commRate(summaryBelford), commAmount(summaryBelford))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Stephern Smith', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summarySSmith), 'Commission Rate': commRate(summarySSmith), 'Earned Commission': totalComm(commRate(summarySSmith), commAmount(summarySSmith))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Tony Barlow', 'Office': 'Fort Worth', 'Commissionable Amount': commAmount(summaryBarlow), 'Commission Rate': manCommRate(summaryBarlow), 'Earned Commission': totalComm(manCommRate(summaryBarlow), commAmount(summaryBarlow))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Von Ysasi', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryYsasi), 'Commission Rate': commRate(summaryYsasi), 'Earned Commission': totalComm(commRate(summaryYsasi), commAmount(summaryYsasi))}, ignore_index=True)
commissionBD = commissionBD.append({'Rep': 'Wanda Coleman', 'Office': 'Austin', 'Commissionable Amount': commAmount(summaryColeman), 'Commission Rate': commRate(summaryColeman), 'Earned Commission': totalComm(commRate(summaryColeman), commAmount(summaryColeman))}, ignore_index=True)
    
austinList = [summaryBrown, summaryPellak, summaryGSmith, summaryDudo, summaryDawson, summaryRifenburg,summaryAudagnotti, summaryAcosta, summaryLaws, summaryBonds, summaryYsasi, summaryColeman]
fw_sa_List = [summaryCanchola, summaryWest, summaryMcCance, summaryVonVogt, summaryAtkins, summaryOverton, summaryBononcini, summaryYokom, summaryDore, summaryFerrigno, summaryHilburn, summaryBelford, summarySSmith, summaryBarlow]

masterList = [austinList, fw_sa_List]
austinDf = pd.DataFrame(columns = ['BID', 'Business', 'Joined', 'Billed', 'Item', 'Payment Date', 'Type', 'Notes', 'Sales', 'Amount', '', 'Total', 'Commission Rate', 'Commission'])
fw_sa_df = pd.DataFrame(columns = ['BID', 'Business', 'Joined', 'Billed', 'Item', 'Payment Date', 'Type', 'Notes', 'Sales', 'Amount', '', 'Total', 'Commission Rate', 'Commission'])

for i in austinList:
    austinDf = austinDf.append(i) 
    if commAmount(i) != 0:
        austinDf = austinDf.append({'Total': commAmount(i), 'Commission Rate': commRate(i), 'Commission': totalComm(commAmount(i), commRate(i))}, ignore_index=True)
    else:
        pass

for i in fw_sa_List:
    fw_sa_df = fw_sa_df.append(i)
    if commAmount(i) != 0:
        fw_sa_df = fw_sa_df.append({'Total': commAmount(i), 'Commission Rate': commRate(i), 'Commission': totalComm(commAmount(i), commRate(i))}, ignore_index=True)
    else:
        pass
    
a = managerTotal('Austin')
b = managerTotal('FwSa')

jessComm = int(a[0]) * .03

austinDf = austinDf.append({'': 'Total', 'Total': a[0], 'Commission Rate': 'Total Commission', 'Commission': a[1]}, ignore_index=True)
austinDf = austinDf.append({'': 'Jessica 3%', 'Total': jessComm}, ignore_index=True)


managerColumns = ['Item', 'Type', 'Notes', 'Processed', 'Sales', 'Item Amt']
austinManager = pd.DataFrame(columns = managerColumns)
fortWorthManager = pd.DataFrame(columns = managerColumns)
sanAntonioManager = pd.DataFrame(columns = managerColumns)

df = summaryCanchola.append(summaryWest)
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
df = df.append(summaryAudagnotti)
df = df.append(summaryAcosta)
df = df.append(summaryLaws)
df = df.append(summaryBononcini)
df = df.append(summaryYokom)
df = df.append(summaryDore)
df = df.append(summaryChevere)
df = df.append(summaryBonds)
df = df.append(summaryWagoner)
df = df.append(summaryFerrigno)
df = df.append(summaryHilburn)
df = df.append(summaryBelford)
df = df.append(summarySSmith)
df = df.append(summaryBarlow)
df = df.append(summaryYsasi)
df = df.append(summaryColeman)

testReport.to_excel("C:/Users/jkreuger/Downloads/GenPayReport.xlsx")

writer = pd.ExcelWriter('C:/Users/jkreuger/Downloads/report test.xlsx', engine='xlsxwriter')
commissionBD.to_excel(writer, sheet_name = 'Summary', index = False)
df.to_excel(writer, sheet_name='New', index = False)
austinDf.to_excel(writer, sheet_name='Austin', index=False)
fw_sa_df.to_excel(writer, sheet_name='Fort Worth.San Antonio', index=False)
writer.save()


driver.quit()
