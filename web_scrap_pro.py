import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
import re
import sqlite3
import traceback

'''Steps  followed: '''
# 1. read urls from file and store valid ones in a list- valid_urls
# 2. Get all words in list - web_text
# 3. iterate through and remove symbols, numbers and white spaces and empty elements- clean_words
# 4. remove common words from csv file
# 5. convert words to lowercase
# 6. find frequency of unique words and store it in a dictionary- word_freq
# 7. sort dictionary word_freq by value
# 8. find density of values and store it in a list
# 9. export 2 columns to excel sheet with url as sheet name
# 10. make chart of top 20 high density words
# 11. Save the data in SQLite table

# read input urls from file and store it in a list- urls
f = open("input_urls", "r")
urls = f.read().splitlines()
valid_urls=[]
workbook = xlsxwriter.Workbook('Web_Analysis.xlsx')
conn = sqlite3.connect('webanalysis.db')
c = conn.cursor()

# iterate through the urls and analyse them
for url in urls:
    try:
        # extract the readable text from the webpage of the url
        req = urllib.request.Request(url, data=None)
        valid_urls.append(url)
        page = urllib.request.urlopen(req)
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
        #remove stop words and common words
        f = open("common_words", "r")
        common_words = f.read().split(",")
        refined_words = [word for word in clean_words if word.lower() not in common_words]
        web_words = []
        for word in refined_words:
            web_words.append(word.lower())
        #get unique words
        web_set=set(web_words)
        unique_words=list(web_set)
        #get frequency of words
        word_freq={}
        for word in unique_words:
            word_freq[word]=web_words.count(word)
        sorted_words = sorted(word_freq, key=word_freq.get, reverse=True)
        #get density of words
        density_words = []
        for r in sorted_words:
            #print(r,word_freq[r])
            density_words.append((word_freq[r]/len(web_words))*100)
        #export data to excel sheet
        sheetname = url.split(".")[1]
        worksheet = workbook.add_worksheet(sheetname)
        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0

        # Iterate over the data and write it out row by row.
        for word in sorted_words:
            worksheet.write(row, col, word)
            worksheet.write(row, col + 1, word_freq[word])
            row += 1
        row = 0
        col = 2
        for word in density_words:
            worksheet.write(row,col,word)
            row +=1
        #Data visualization using excel line chart in Web_Analysis.xlsx workbook
        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({
            'values': '=' + sheetname + '!$A$1:$A$10',

        })
        chart.add_series({
            'values': '=' + sheetname + '!$C$1:$C$10',

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
        # for row in cursor:
        #     print(row[0],row[1],row[2], "\n")

        print(url,"- ok")
    except Exception as e:
        print("Exception:", e)
        #traceback.print_exc()
#close connections
workbook.close()
conn.close()


