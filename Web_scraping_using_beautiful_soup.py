import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
import re
import sqlite3
import traceback

#getting the URL list
f = open("input_urls.txt", "r")
valid_urls= f.read().splitlines()

#getting the search words from the user
while True:
    try:
        search_words=input("Enter the words you are looking for separated by ,:").split(",")
        search_words=list(search_words)
        search_words=[x.lower() for x in search_words]
        search_words=[re.sub(r'[^a-zA-Z]', r'', x) for x in search_words]
        break
    except Exception as e:
        print(e,"invalid input,Enter the words you are looking for separated by spaces:")

workbook = xlsxwriter.Workbook('Web_Analysis.xlsx')
conn = sqlite3.connect('webanalysis.db')
c = conn.cursor()


for url in valid_urls:
    page = urllib.request.urlopen(url)
    soup = BeautifulSoup(page, "html.parser")
    for script in soup(["script", "style", "head", "title", "meta", "[document]"]):
        script.extract()
    text = soup.get_text()
    web_text = [line.strip() for line in text.split()]

    #remove numbers, special symbols, spaces
    clean_words = []
    for text in web_text:
        text = re.sub(r'[^a-zA-Z]', r'', text)
        clean_words.append(text)
    #remove None, 0
    clean_words = list(filter(None, clean_words))

    #changing all the words to lower case    
    web_words = []
    for word in clean_words:
        web_words.append(word.lower())
    #making all a string with all the words separated with space
    web_words=" ".join(web_words)
    #get frequency of words
    word_freq={}
    for word in search_words:
        word_freq[word]=len(re.findall(word,web_words))
    sorted_words = sorted(word_freq, key=word_freq.get, reverse=True)
    #get density of words
    density_words = []
    for r in sorted_words:
        density_words.append((word_freq[r]/len(web_words))*100)
    #export data to excel sheet
    sheetname =url.split(".")[1]

    worksheet = workbook.add_worksheet(sheetname)
    # Start from the first cell. Rows and columns are origin indexed.
    row = 1
    col = 0
    worksheet.write("A1","Word")
    worksheet.write("B1","Count")
    worksheet.write("C1","Density")

    # Iterate over the data and write it out row by row.
    for word in sorted_words:
        worksheet.write(row, col, word)
        worksheet.write(row, col + 1, word_freq[word])
        row += 1
    row = 1
    col = 2
    for word in density_words:
        worksheet.write(row,col,word)
        row +=1
    #Data visualization using excel line chart in Web_Analysis.xlsx workbook
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        'values': '=' + sheetname + '!$A$2:$A$11',

                })
    chart.add_series({
        'values': '=' + sheetname + '!$C$2:$C$11',

                })

    chart.set_legend({'position': 'none'})
    # Add a chart title and some axis labels.
    chart.set_title({'name': 'Results of Web Scraping'})
    chart.set_y_axis({'name': 'Word Density'})
    chart.set_x_axis({'name': 'Sno of Words'})

    worksheet.insert_chart('F5', chart)

    #store data in webanalysis.db using SQLite
    table_list=[]

    for word in sorted_words:
        table_list.append([word,word_freq[word],(word_freq[word]/len(web_words))*100])
        tablename=sheetname
    # Create table
    c.execute("create table if not exists %s (word text, frequency real, density real)" % (tablename))
    for word in table_list:
        c.execute("insert into %s values(?,?,?)" % (tablename), word)

    # Save (commit) the changes
    conn.commit()

    cursor = conn.execute("select * from %s" % (tablename))
    print(url,"- ok")


workbook.close()
conn.close()

