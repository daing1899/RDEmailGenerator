import os
import pandas as pd
import numpy as np

def readMappings(fileName):
    mappings = readDataFromSpreadSheet(fileName, 'Mappings')
    for rowIndex in range(len(mappings['columnType'])):
        if (mappings['columnType'][rowIndex] == 'GivenName'):
            givenNameKey = mappings['columnName'][rowIndex]
        if (mappings['columnType'][rowIndex] == 'PartnerName'):
            partnerNameKey = mappings['columnName'][rowIndex]
        if (mappings['columnType'][rowIndex] == 'Address'):
            addressKey = mappings['columnName'][rowIndex]
    return givenNameKey, partnerNameKey, addressKey


def readDataFromSpreadSheet(fileName, sheetName):
    if fileName=='':
        fileName ='default'
    
    fullFilename = fileName + '.xlsx'
    if not os.path.isfile(fullFilename):
        return
    read_df = pd.read_excel(fullFilename, sheet_name=sheetName)
    print(read_df.columns)
    return read_df

def findOrderCodeIndex(df, orderCode):
    for rowIndex in range(len(df['Order code'])):
        if (df['Order code'][rowIndex] == orderCode):
            return rowIndex
    return 

def calculateDish(df, groupNumber):
    peoplePerGroup = 3 + 1
    groupAmount = (df.shape[0]/peoplePerGroup)
    for i in range(int(groupAmount)):
        lineIndex = peoplePerGroup*i
        # print(f'line0: {df.iloc[lineIndex, 0]}')
        # print(f'line1: {df.iloc[peoplePerGroup*i+1, 1]}')
        # print(f'line2: {df.iloc[peoplePerGroup*i+1, 2]}')
        if (int(df.iloc[lineIndex+1, 0]) == groupNumber):
            dishType = df.iloc[lineIndex, 0] 
            return translateDish(0) + cleanDishType(dishType)
        elif (int(df.iloc[peoplePerGroup*i+1, 1]) == groupNumber): 
            dishType = df.iloc[peoplePerGroup*i, 1]
            return translateDish(1) + cleanDishType(dishType)
        elif (int(df.iloc[peoplePerGroup*i+1, 2]) == groupNumber): 
            dishType = df.iloc[peoplePerGroup*i, 2]
            return translateDish(2) + cleanDishType(dishType)
    return -1

def translateDish(code):
    if (code == 0):
        return 'Starter'
    elif (code == 1):
        return 'Main Dish'
    elif (code == 2):
        return 'Dessert'
    return 'Error'

def cleanDishType(dishType):
    if (dishType == ' '):
        return ''
    else:
        return ' ' + dishType

def calculateAddress(groupData, pretixData, groupIndex):
    order1 = groupData['Key1'][groupIndex-1]
    registerIndex = findOrderCodeIndex(pretixData, order1)
    address = pretixData[addressKey][registerIndex]
    splitAddress = address.splitlines()
    trimmedAddress = " ".join(splitAddress)
    return trimmedAddress

def calculateNames(groupData, pretixData, groupIndex):
    order1 = groupData['Key1'][groupIndex-1]
    registerIndex1 = findOrderCodeIndex(pretixData, order1)
    name1 = pretixData[givenNameKey][registerIndex1]

    order2 = groupData['Key2'][groupIndex-1]
    registerIndex2 = findOrderCodeIndex(pretixData, order2)
    if (registerIndex2 is not None):
        name2 = pretixData[givenNameKey][registerIndex2]
    else:
        name2 = pretixData[partnerNameKey][registerIndex1]
        if (name2 is np.nan):
            name2 = groupData['Person2'][groupIndex-1]
    return name1, name2

def calculateEmails(groupData, pretixData, groupIndex):
    order1 = groupData['Key1'][groupIndex-1]
    registerIndex1 = findOrderCodeIndex(pretixData, order1)
    email1 = pretixData['E-mail'][registerIndex1]

    order2 = groupData['Key2'][groupIndex-1]
    registerIndex2 = findOrderCodeIndex(pretixData, order2)
    if (registerIndex2 is not None):
        email2 = pretixData['E-mail'][registerIndex2]
    else:
        email2 = ""
    return email1, email2

def calculateAddressRoute(pretixData, groupData, routeData, groupNumber):
    peoplePerGroup = 3 + 1
    groupAmount = routeData.shape[0]
    addressArray = ['', '', '']
    peoplePerGroup = 3 + 1
    routeSize = len(routeData.columns)
    for dish in range(int(routeSize)):
        for lineIndex in range(int(groupAmount)):
            # lineIndex = peoplePerGroup*lineIndex
            if (lineIndex%peoplePerGroup != 0):
                # print(f'line to search:{routeData.iloc[lineIndex, dish]}')
                if (routeData.iloc[lineIndex, dish] == groupNumber): 

                    homeGroup = routeData.iloc[int(lineIndex/peoplePerGroup) * peoplePerGroup +1, dish] 
                    print(f'home group: {homeGroup}')
                    # order1 = groupData['Key1'][homeGroup]
                    # registerIndex = findOrderCodeIndex(pretixData, order1)
                    print (f'{calculateAddress(groupData, pretixData, homeGroup)}')
                    addressArray[dish] = calculateAddress(groupData, pretixData, homeGroup)
                    break
        print(f'finished dish {dish} for {groupNumber}') 
    return addressArray

def createEmail(folder, groupIndex, emails, groupNames, dish, addressRoute):
    f = open(folder + "Email.md", "r")
    text = f.read()
    f.close()
    text = emails[0] + '\n' + emails[1] + '\n' + text

    text = text.replace("{1}", str(groupNames[0]))
    text = text.replace("{2}", str(groupNames[1]))

    text = text.replace("{3}", str(dish))

    str4 = ""
    str6 = ""
    str8 = ""

    if (dish == 'Starter'):
        str4 = '(Your place)'
    elif (dish == 'Main Dish'):
        str6 = '(Your place)'
    elif (dish == 'Dessert'):
        str8 = '(Your place)'
    
    text = text.replace("{4}", str4)
    text = text.replace("{6}", str6)
    text = text.replace("{8}", str8)

    text = text.replace("{5}", addressRoute[0])
    text = text.replace("{7}", addressRoute[1])
    text = text.replace("{9}", addressRoute[2])


    f2 = open(folder + "Email" + str(groupIndex) + ".md", "w")
    f2.write(text)
    f2.close()

# givenNameKey = 'Attendee name: Given name'
# partnerNameKey = 'My team partners\'s name:'
# addressKey = 'My address, please also specify where to ring the Bell (e.g. WG 5. Stock, or whatever is written on your bell)'


folder = 'RC 24-02/'
sheetName = 'Check-in list'
fileName = folder + 'PretixData' 
#readData
pretixData = readDataFromSpreadSheet(fileName, sheetName)
#readGroupsData
groupFileName = folder + 'GroupData'
groupSheet = 'Groups'
groupData = readDataFromSpreadSheet(groupFileName, groupSheet)
routeFileName = folder + 'GroupData'
routeSheet = 'Route'
routeData = readDataFromSpreadSheet(routeFileName, routeSheet)

givenNameKey, partnerNameKey, addressKey = readMappings(groupFileName)
#readTemplate
#number mapping to property

f = open("demofile2.txt", "w")

for groupIndex in range(1,len(groupData['Key1'])+1):
    address = calculateAddress(groupData, pretixData, groupIndex)
    groupNames = calculateNames(groupData, pretixData, groupIndex)
    emails = calculateEmails(groupData, pretixData, groupIndex)
    dish = calculateDish(routeData, groupIndex)
    print(dish)
    addressRoute = calculateAddressRoute(pretixData, groupData, routeData, groupIndex)
    print(addressRoute)
    f.writelines (f'Group {groupIndex}\n')

    f.writelines (f'Emails\n')
    for i in range(len(emails)):
        f.write(str(emails[i]) + '\n')

    f.writelines (f'Group Names\n')
    for i in range(len(groupNames)):
        f.write(str(groupNames[i]) + '\n')

    f.write(f'Dish {dish}\n')
    for i in range(len(addressRoute)):
        f.write(str(addressRoute[i]) + '\n')

    createEmail(folder, groupIndex, emails, groupNames, dish, addressRoute)
f.close()