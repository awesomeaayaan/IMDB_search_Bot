from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import re

browser = Selenium(auto_close=False)
excel = Files()


def open_browser():
    url = 'https://www.imdb.com/'
    browser.open_available_browser(url, maximized=True)
def read_from_excel():
    excel.open_workbook('movies.xlsx')
    movies_table = excel.read_worksheet_as_table(header=True)
    print(movies_table)
    excel.close_workbook()
    return movies_table

def fetches_movie_details(movie):
    movie_name = movie['Movie']
    browser.input_text('//*[@id="suggestion-search"]', movie_name)
    browser.click_button('suggestion-search-button')
    browser.click_element('//*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')
    text = browser.get_text('//*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul')
    input_movie = movie
    lines = text.split('\n')
    #print(lines)
    

    # Convert list to dictionary
    movies_dict = {}
    for i in range(0, len(lines), 2):
        movie = lines[i]
        year = lines[i + 1]
        year_numeric = re.findall(r'\b\d+\b', year)  # Extract numerical year using regex
        if year_numeric:
            movies_dict[movie] = int(year_numeric[0])

    input_movie = movie_name  # Replace with your input movie

    matching_movies = {}
    for movie, year in movies_dict.items():
        if movie == input_movie:
            matching_movies[movie] = year

    if matching_movies:
        latest_year = max(matching_movies.values())
        latest_movie = max(matching_movies, key=matching_movies.get)
        index = list(movies_dict.keys()).index(latest_movie)
        print(f"Latest matching movie: {latest_movie}")
        print(f"Year: {latest_year}")
        print(f"Index: {index}")
    else:
        print("No matching movie found.")

    browser.click_element(f'//*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[{index}]/div[2]')
    # except:
    #     print("Not found")

     
def main():
    #open_browser()
    #create_table()
    open_browser()
    movies_table = read_from_excel()
    for movie in movies_table:
        fetches_movie_details(movie)
    

if __name__ == "__main__":
    main()
