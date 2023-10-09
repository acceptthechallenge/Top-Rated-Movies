from bs4 import BeautifulSoup
import requests
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
# print(excel.sheetnames)
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating'])

try:

    HEADERS = {
        'User-Agent': 'Mozilla/5.0 (iPad; CPU OS 12_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148'}

    # This will fetch the data from IMDB
    source = requests.get('https://www.imdb.com/chart/top/', headers=HEADERS)

    # It will capture the error if in case the website isn't responding
    source.raise_for_status()

    soup = BeautifulSoup(source.text, 'html.parser')

    movies = soup.find(
        'ul', class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-3f13560f-0 sTTRj compact-list-view ipc-metadata-list--base").find_all('li', class_="ipc-metadata-list-summary-item sc-59b6048d-0 jemTre cli-parent")

    for movie in movies:

        name = movie.find(
            'div', class_="ipc-metadata-list-summary-item__c").a.get_text(strip=True).split('.')[1]

        rank = movie.find(
            'div', class_="ipc-metadata-list-summary-item__c").a.get_text(strip=True).split('.')[0]

        year = movie.find(
            'span', class_="sc-6fa21551-8 bnyjtW cli-title-metadata-item").text

        rating = movie.find(
            'span', class_="ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating").text.strip(' ')[0:3]

        print(rank, name, year, rating)

        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)

excel.save('IMDB Movie Ratings.xlsx')
