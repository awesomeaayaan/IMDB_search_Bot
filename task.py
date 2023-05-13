from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import logging
import time
import re
import sqlite3



browser_lib = Selenium(auto_close=False)
FILE_NAME = "movies.xlsx"

url = 'https://www.imdb.com/'
excel = Files()
logging.basicConfig(level=logging.INFO,format='%(levelname)s:%(message)s')
#path

ratingpath = 'xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[2]/div/div[1]/a/span/div/div[2]/div[1]/span[1]'
storylinepath = 'xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/p'
taglinepath1 = 'xpath://*[@id="__next"]/main/div/section[1]/div/section/div/div[1]/section[6]/div[2]/ul[2]/li[1]/div/div/ul'
taglinepath2 = 'xpath://*[@id="__next"]/main/div/section[1]/div/section/div/div[1]/section[7]/div[2]/ul[2]/li[1]/div/div/ul'
genrespath = 'xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/div[1]/div[2]'
user_review_path = 'xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[1]/div/div[2]/ul/li[2]/a'
review1path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[1]/div/div[1]/a'
review2path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[2]/div[1]/div[1]/a'
review3path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[3]/div[1]/div[1]/a'
review4path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[4]/div[1]/div[1]/a'
review5path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[5]/div[1]/div[1]/a'
con = sqlite3.connect("movie.db")
cur = con.cursor()
def open_browser():
    # use headless=True so robot work without opening browser, removi maximize argument
    # maximize help robot to open browser in full screen
    browser_lib.open_available_browser(url, maximized=True)


def create_table():
    try:
        cur.execute("CREATE TABLE movie(Mid INTEGER PRIMARY KEY AUTOINCREMENT, movie_name TEXT , tagline TEXT, storyline TEXT, rating TEXT, genres TEXT, review_1 TEXT, review_2 TEXT, review_3 TEXT, review_4 TEXT, review_5 TEXT, status TEXT)")

    except:
        print("Table already exist")



def read_from_excel():
    excel.open_workbook('movies.xlsx')
    movies_table = excel.read_worksheet_as_table(header=True)
    excel.close_workbook()
    return movies_table
    
def insert_movies(movie):
    movie_name = movie['Movie']
    browser_lib.input_text('//*[@id="suggestion-search"]',movie_name)
    browser_lib.click_button('suggestion-search-button')
    filter_movie = 'Movies'
    get_all = browser_lib.get_text('//*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')

    if get_all == filter_movie:
        browser_lib.click_element('//*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')
    exact_movie_index = search_for_exact_match(movie_name)
    search_for_latest_movie_and_click_the_movie(exact_movie_index, movie_name)
def extract_data(movie_name):
    status = "Success"
    try:
        rating = browser_lib.get_text(ratingpath)
    except:
        rating = 'No rating'
    try:
        storyline = browser_lib.get_text(storylinepath)
    except:
        storyline = 'no storyline'
    try:
        genres = browser_lib.get_text(genrespath)
    except:
        genres = 'no genres'
    #print(rating , '*****',genres,'********',storyline)
    browser_lib.execute_javascript("window.scrollTo(0,3500)")
    time.sleep(2)
    try:
        tagline = browser_lib.get_text(taglinepath1)
    except:
        try:
            tagline = browser_lib.get_text(taglinepath2)
        except:
            tagline = "--N/A--"
    browser_lib.execute_javascript("window.scrollTo(3500,0)")
    time.sleep(2)
    #print(tagline)
    #tagline = browser_lib.get_text(taglinepath1)
    browser_lib.click_element(user_review_path)
    try:
        review1 = browser_lib.get_text(review1path)
        rev1 = remove_punctuations(review1)
        print(rev1)
    except:
        print('review not find')
    

    try:
        review2 = browser_lib.get_text(review2path)
        rev2 = remove_punctuations(review2)
        print(rev2)
    except:
        print('review not find')

    try:
        review3 = browser_lib.get_text(review3path)
        rev3 = remove_punctuations(review3)
        print(rev3)
    except:
        print('review not found')
    try:
        review4 = browser_lib.get_text(review4path)
        rev4 = remove_punctuations(review4)
        print(rev4)
    except:
        print('review not found')
    try:
        review5 = browser_lib.get_text(review5path)
        rev5 = remove_punctuations(review5)
        print(rev5)
    except:
        print('review not found')

    movie_data = {
        "movie_name": movie_name,
        "storyline": storyline,
        "rating": rating,
        "tagline": tagline,
        "genres": genres,
        "review1": rev1,
        "review2": rev2,
        "review3": rev3,
        "review4": rev4,
        "review5": rev5,
        "status": status
    }
    insert_into_table(movie_data)
def search_for_latest_movie_and_click_the_movie(exact_movie_index, movie_name):
    latest_release_year = 0
    latest_release_year_index = 0
    j = 0
    for j in exact_movie_index:
        j = str(j)

        movie_year = browser_lib.get_text(
            'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + j + ']/div[2]/div/ul[1]/li')
        movie_year = int(movie_year)
        if movie_year > latest_release_year:
            latest_release_year = movie_year
            latest_release_year_index = j
    latest_release_year_index = str(latest_release_year_index)
    try:

        browser_lib.click_element(
            'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + latest_release_year_index + ']/div[2]/div/a')
        extract_data(movie_name)
    except:
        status = "No exact match found"
        no_movie_found(movie_name,status)

def no_movie_found(movie_name, status):
    rating = "--N/A--"
    storyline = "--N/A--"
    tagline = "--N/A--"
    genres = "--N/A--"
    review1 = "--N/A--"
    review2 = "--N/A--"
    review3 = "--N/A--"
    review4 = "--N/A--"
    review5 = "--N/A--"
    movie_data = {
        "movie_name": movie_name,
        "storyline": storyline,
        "rating": rating,
        "tagline": tagline,
        "genres": genres,
        "review1": review1,
        "review2": review2,
        "review3": review3,
        "review4": review4,
        "review5": review5,
        "status": status
    }
    insert_into_table(movie_data)
    
def search_for_exact_match(movie_name):
    search_result_count = browser_lib.get_element_count(
        'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li')

    if search_result_count > 5:
        search_result_count = 5
    search_result_count = search_result_count + 1

    exact_movie_index = []

    for i in range(1, search_result_count):

        i = str(i)
        movie_title = browser_lib.get_text(
            'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + i + ']/div[2]/div/a')
        movie_title = str(movie_title)
        movie_title = movie_title.lower()
        movie_title = remove_punctuations(movie_title)
        movie_title = movie_title.strip()

        movie_name = str(movie_name)
        movie_name = movie_name.lower()
        movie_name = remove_punctuations(movie_name)
        movie_name = movie_name.strip()

        if movie_title == movie_name:
            exact_movie_index.append(i)
        
    return exact_movie_index

def insert_into_table(movie_data):
    insert_sql = """INSERT INTO movie(movie_name, 
                            tagline, 
                            storyline,
                            rating,
                            genres,
                            review_1,
                            review_2,
                            review_3,
                            review_4,
                            review_5,
                            status
                            ) 
                            VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """
    cur.execute(insert_sql, (movie_data["movie_name"],
                             movie_data["tagline"],
                             movie_data["storyline"],
                             movie_data["rating"],
                             movie_data["genres"],
                             movie_data["review1"],
                             movie_data["review2"],
                             movie_data["review3"],
                             movie_data["review4"],
                             movie_data["review5"],
                             movie_data['status']
                             
                             ))
    con.commit()
    # data = cur.execute("Select * from movie")
    # logging.info(data.fetchall())
        

def remove_punctuations(string):
    
    pattern = r'[\"\',]'
    return re.sub(pattern, '', string)
      
 
def main():
    open_browser()
    create_table()
    movies_table = read_from_excel()
    for movie in movies_table:
        insert_movies(movie)
    



if __name__ == "__main__":
    main()
