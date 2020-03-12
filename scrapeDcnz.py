'''
@author: jorge
'''


from bs4 import BeautifulSoup
import xlsxwriter
import requests


for x in range(6):

    # original page
    page = 'https://dcnz.org.nz/practitioners/PractitionerSearchForm?Practice=&Surname=&Name=&Address=&PersonID=&action_doPractitionerSearch=Search&g-recaptcha-response=03AERD8Xqa_hdvo9S9S2tfaSUTz4Y8CNcAKCRX6y9zx50LYv768WYADW2VdjsICEdu0EYuhZRUcL1q6BYn4CTOV4yS4ZQ8bOEtprv0bUpMnH0-fvqT09uvM4ux1P-V0YoO1ElkW8erLVZ9Iuym1Otcn9lOk9qEmSHwCfKWNnmqDjt2dzDRNulaPgXlemjAX1QUMLMFs_BGWNsZ0cpGQggpU4w8sgohd1rclW5-pF4QbpKmgQryQsrcJTcusKhTOea1KQhhVHwUnxt3BPXFQvVqxcZUtIWKcFhXn_6mI5pN3aoWm0VWMTy0gz35gmNSZOHXu_tTAYHMY6KEaxTJBlP9HVWXCjghgFtrKIrBxaejAto8mEN5m_UudYMuEF6zWQDbdboJ1IVOjuHU&start={}'
    pageAux = page.format(x*10)
    # first list with practitioners
    r = requests.get(pageAux)

    if r.status_code == 200:
        
        # Create variable to save html 
        soup = BeautifulSoup(r.content, 'html.parser')
    
        # create excell and add a sheet
        workbook = xlsxwriter.Workbook('practitioner.xlsx')
        worksheet = workbook.add_worksheet() 
        
        # characteristics before name of type of practitioner
        characteristicsBeforePract = {'Practice': 1, 'Name': 2}
        # position to add a new characteristic in characteristicsBeforePract
        positionBeforePract = 3 
        # characteristics after name of type of practitioner
        characteristicsAfterPract = {}
        
        # boolean variable to check if the type of practitioner characteristic was found 
        practiceFound = False
        
        # list with all the practitioners
        listPractitioners = []
        
        
        for table in soup.find_all('table', class_ = 'practitioner-result'):
            
            # create a new practitioner
            practitioner = dict({})
            
            #find name of the practitioner
            name = table.find_previous_sibling('h3').text
            # add name of the new practitioner
            practitioner['Name'] = name.replace('-','').rstrip()
            # characteristics
            for child in table.findChildren('tr'):
                if (child.find('td')):
                    key = child.th.text.replace(':', '')
                    practitioner[key] = child.td.text
                    if practiceFound == False:
                        # characteristic before practitioner
                        if not(key in characteristicsBeforePract):
                            characteristicsBeforePract[key]= positionBeforePract
                            positionBeforePract += 1
                    else: #characteristic after practitioner
                        if not(key in characteristicsAfterPract):
                            characteristicsAfterPract[key]= 0
                else: # type of practitioner characteristic is found
                    practitioner['Practice'] = child.th.text
                    practiceFound = True
            # add new practitioner        
            listPractitioners.append(practitioner)  
            practiceFound = False
        
        lastKey = 0
        # print in the excell characteristics before type of practitioner
        for key, value in characteristicsBeforePract.items():
            worksheet.write(value, 0, key)
            lastKey = value
        # print in the excell characteristics after type of practitioner        
        for key, value in characteristicsAfterPract.items():
            lastKey += 1
            worksheet.write(lastKey, 0, key)
            characteristicsAfterPract[key] = lastKey
        
        column = 1
        # print practitioners using characteristicsBeforePract and characteristicsAfterPract to 
        # get the number of row where the characteristic is
        for practitioner in listPractitioners:
            for key, value in practitioner.items():
                if key in characteristicsBeforePract:
                    worksheet.write(characteristicsBeforePract[key], column, value)
                if key in characteristicsAfterPract:
                    worksheet.write(characteristicsAfterPract[key], column, value)
            column += 1
            
        workbook.close()
