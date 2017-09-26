import csv, os
from datetime import datetime


Super_Duper = os.path.expanduser('~\Documents\SUPER DUPER')
if not os.path.exists(Super_Duper):
    os.makedirs(Super_Duper)


class NewCsvToOldCsv():

    def __init__(self, newCsvFile):

        self.newCsvFile = newCsvFile
        self.ISIN = None

        # get script location
        script_dir = os.path.dirname(__file__)
        temp_path = script_dir + '\CSVMappings\OldCsvMap.txt'

        # get New Csv to Old Csv map
        with open(temp_path, 'r') as f:
            self.oldCsvMap = f.readlines()
        self.oldCsvMap = [x.strip().split('\t') for x in self.oldCsvMap]

    def readNewCsv(self):

        with open(self.newCsvFile, newline='') as f:
            reader = csv.reader(f,
                                delimiter='|',
                                escapechar='',
                                quotechar="'")
            n = 0
            for row in reader:
                # we need data from second row
                if n == 1:
                    self.newCsvArray = row
                # sometimes csv from fondsportalen has extra rows
                # all of these rows should be added to the end of second row
                elif n > 1:
                    if row:
                        # remove any ' symbols in file
                        self.newCsvArray[-1] += ' ' + row[0].replace("'","")
                        self.newCsvArray += row[1:]
                n += 1

        self.ISIN = self.newCsvArray[0]


    def rearangeValues(self):

        mainTags = ['Type', 'env', 'Status']
        extraTags = ['Termin',
                    'Foerste_emissionsdato',
                    'Sidste_emissionsdato',
                    'Traekningsberegningsdato',
                    'Traekningsdato',
                    'Publiceringsdato',
                    'Terminsda']

        mainText = ['B', '', '34']
        extraText = []

        # fields with dates
        dates = [13, 37, 38, 39, 40, 41, 55]
        # fields with options yes or no
        yesNo = [19, 20, 23]

        # add emtpy fields to mainText array
        for i in range(67):
            mainText.append('')

        # adjustments necessary to adapt new csv to old
        for idx, elem in enumerate(self.oldCsvMap):
            mainTags.append(elem[0])
            if (elem[1] != 'None'):
                # first interest payment year
                if idx == 42:
                    mainText[idx+3] = ((self.newCsvArray[int(elem[1])])[:4]).strip()
                # adjust format of date fields
                elif idx in dates:
                    d = (self.newCsvArray[int(elem[1])]).strip()
                    if str(d) != '0':
                        mainText[idx+3] = '-'.join([d[:4], d[4:6], d[6:]])
                    elif int(d) == 0:
                        mainText[idx+3] = ''
                # adjust values of Yes or No fields
                elif idx in yesNo:
                    value = (self.newCsvArray[int(elem[1])]).strip()
                    if value == 'Y':
                        mainText[idx+3] = '1'
                    elif value == 'N':
                        mainText[idx+3] = '2'
                    else:
                        mainText[idx+3] = value
                # all below ifs are for fields with dropdown options
                elif idx == 11:
                    value = (self.newCsvArray[int(elem[1])]).strip()
                    if value == '1':
                        mainText[idx+3] = '1'
                    elif value == '2':
                        mainText[idx+3] = '4'
                    elif value == '4':
                        mainText[idx+3] = '3'
                    else:
                        value == ''
                elif idx == 14:
                    value = (self.newCsvArray[int(elem[1])]).strip()
                    if value == '1':
                        mainText[idx+3] = '1'
                    elif value == '4':
                        mainText[idx+3] = '2'
                    else:
                        mainText[idx+3] = ''
                elif idx == 15:
                    value = (self.newCsvArray[int(elem[1])]).strip()
                    if value in ['4','5','8','9']:
                        mainText[idx+3] = ''
                    else:
                        mainText[idx+3] = value
                elif idx == 27:
                    mainText[idx+3] = '3'
                elif idx == 32:
                    value = (self.newCsvArray[int(elem[1])]).strip()
                    mainText[idx+3] = str(int(float(value)))
                elif idx == 43:
                    value = (self.newCsvArray[int(elem[1])]).strip()
                    if value == '4':
                        mainText[idx+3] = '4'
                    else:
                        mainText[idx+3] = ''
                elif idx == 60:
                    value = (self.newCsvArray[int(elem[1])]).strip()
                    if value == '3':
                        mainText[idx+3] = '3'
                    elif value == '1':
                        mainText[idx+3] = '1'
                    else:
                        mainText[idx+3] = ''
                else:
                    mainText[idx+3] = (self.newCsvArray[int(elem[1])]).strip()

        # add opening data if there is any
        n = 88
        x = 1
        while True:
            if self.newCsvArray[n] != '':
                tempArray = [''] * 77
                tempArray[0] = 'O'
                tempArray[70] = '%d'%x
                for i in range(6):
                    # print (i)
                    d = self.newCsvArray[n+i]
                    tempArray[71+i] = '-'.join([d[:4], d[4:6], d[6:]])
                extraText.append(tempArray[:])
                n += 6
                x += 1
            else:
                break

        if extraText:
            mainTags += extraTags

        return mainTags, mainText, extraText

    def csvGenerator(self):

        mainTags, mainText, extraText = self.rearangeValues()

        nameDate = str(datetime.today().strftime('%Y%m%d'))

        with open((Super_Duper + '\\' + str(self.ISIN) + '-' + nameDate + '.csv'), 
                  'w', 
                  newline='') as outfile:

            writer = csv.writer(outfile, 
                                quoting=csv.QUOTE_NONE, 
                                delimiter='|', 
                                escapechar='', 
                                quotechar="'"
                                )
            writer.writerow(mainTags)
            writer.writerow(mainText)
            if extraText:
                for row in extraText:
                    writer.writerow(row)

# test = NewCsvToOldCsv('Kladde til NASDAQ OMX Copenhagen-vp.csv')
# test = NewCsvToOldCsv('test.csv')
# test.readNewCsv()
# test.csvGenerator()