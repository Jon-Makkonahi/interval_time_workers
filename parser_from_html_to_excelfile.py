import pandas
import requests
import bs4


url = ''


source = requests.get(url).text

soup = bs4.BeautifulSoup(source, 'lxml')

table = soup.select_one('name_classes')
historical_data = pandas.read_html(str(table))
historical_data = historical_data[0]

historical_data.to_excel('download_table.xlsx',index=False)
