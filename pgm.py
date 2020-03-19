from urllib.request import Request
from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter
import re

com_chars = ['more','has','it','in','by','the','ies','and','be','these','not','such',
             'can','then','when','which','one','of','as','from','ed','ing','s','on',
             'that','1','2','3','4','5','6','7','8','9','0','was','a','be','ly','is',
             'with','e','are','for','an','ia','or','to']
frequent_words = []
frequency = []
density = []
mass = []
data = []
dct = {}
heading = ['WORDS','FREQUENCY','DENSITY']

try:
    input_file = open(input('Enter your input file name with extension : '))
    r_file = input_file.read()
except Exception as e:
    print(type(e),':Incorrect')
else:
    print('Input file read successfully')

url = r_file.split()
print('Received URL count is ',len(url))

xl_file = xlsxwriter.Workbook('results.xlsx')
print('Please wait...')

count = 0
while count < len(url):
    request = Request(url [count],data = None)
    page = urlopen(request)
    soup = BeautifulSoup(page,'html.parser')
    Heading = [soup.title.string]
    for script in soup(['script','style']):
        script.extract()
    text = soup.get_text().lower()
    fltr = filter(None,re.split(r'\W|d',text))
    dct.clear()
    word_count = len(text)
    for word in fltr:
        word = word.lower()
        if word in com_chars:
            continue
        if word not in dct:dct[word] = 1
        else:
            dct[word] += 1
    srt = sorted(dct.items(),key = lambda v: v [1],reverse = True,) [:5]
    density.clear()
    for sk,sv in srt:
        key = len(sk)
        den = (key / word_count) * 100
        density.append(den)
    var = [(k,v) for k,v in srt]
    data.clear()
    for r in var:
        data.append(r)
    frequent_words.clear()
    frequency.clear()
    for k,v in data:
        frequent_words.append(k)
        frequency.append(v)
    mass.clear()
    for sv in density:
        mass.append(sv)

    xl_sheet = xl_file.add_worksheet()
    style = xl_file.add_format({'bold': 1})
    column = [frequent_words,frequency,mass]
    xl_sheet.write_row('D6', heading, style)
    xl_sheet.write_column('D7',column[0])
    xl_sheet.write_column('E7',column[1])
    xl_sheet.write_column('F7',column[2])
    graph = xl_file.add_chart({'type': 'column'})
    graph.add_series({'values': '=Sheet1!$E$7:$E$11',
                      'categories':'=Sheet1!$D$7:$D$11',
                      'trendline':False})
    graph.set_x_axis({'name':'Words',
                      'name_font': {'size': 14, 'bold': True}})
    graph.set_y_axis({'name':'Frequency',
                      'name_font': {'size': 14, 'bold': True},
                      'major_gridlines': {'visible': False}})
    xl_sheet.insert_chart('I5',graph)
    count += 1

print('Result generated successfully in excel file')
xl_file.close()


