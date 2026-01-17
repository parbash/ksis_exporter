import requests
import re
import os
import csv
from bs4 import BeautifulSoup

prop_id = input("Enter comp prop_id (ex: 8819): ")
url = f"https://ksis.eu/resultx.php?id_prop={prop_id}"

try:
    response = requests.get(url)
    response.raise_for_status() 
    html_content = response.text
except requests.exceptions.RequestException as e:
    print(f"Error fetching the URL: {e}")
    exit()

sessions = []
all_data = []

soup = BeautifulSoup(html_content, "html.parser")
comp_name = soup.find('h3').get_text().replace("/","-")

# Get list of sessions
print('Parsing sessions...')
select_tag = soup.find('select', id='id_sut')
if select_tag:
    options = select_tag.find_all('option')
    for option in options:
        value = option.get('value')
        content = option.get_text()
        print(f"    Session: {content}")
        result_url = f"https://ksis.eu/load_result_total_ksismg_art.php?lang=en&id_prop={prop_id}&id_sut={value}&rn=null&mn=null&state=-1&age_group=&award=-1&nacinie=undefined"
        try:
            results = requests.get(result_url)
            results.raise_for_status() 
            result_content = results.text
        except requests.exceptions.RequestException as e:
            print(f"Error fetching the URL: {e}")
            exit()
        soup2 = BeautifulSoup(result_content, "html.parser")
        table = soup2.find('table', attrs={'id': 'myTablePrihlasky'})
        data = []
        # Parse Data from table
        for row in table.find_all('tr'):
            cells = row.find_all(['td'])
            row_data = [str(cell) for cell in cells]
            if row_data:
                row_ath = row_data[2]
                row_name, row_club = row_ath.split('<br/>')
                row_name = row_name.split('>')[2].replace("</a","")
                row_club = row_club.replace('</td>',"").rstrip()
                row_yob = re.sub(r"<.*?>", "", row_data[3] )
                row_score = re.sub(r"<.*?>", "", row_data[8] )
                data.append([row_name,row_club,row_yob,row_score])
                all_data.append([content.lstrip(),row_name,row_club,row_yob,row_score])
    # Write CSV File            
    file_path = f'{comp_name}.csv'
    with open(file_path, "w", newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter=',')
        writer.writerows(all_data)
    print(f"\n{file_path} is ready.")

else:
    print("Select tag not found. Wrong prod_id or the website structure has changed.\n Exiting...")


