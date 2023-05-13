from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files

import time
import sqlite3

browser = Selenium()
excel = Files()


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


def open_the_browser():
    browser.open_available_browser("https://imdb.com/")


def connect_database():
    con = sqlite3.connect("movie.db")
    return con


def create_cursor(con):
    cur = con.cursor()
    return cur


def create_table(cur):
    try:
        cur.execute("CREATE TABLE movie(Mid INTEGER PRIMARY KEY AUTOINCREMENT, movie_name TEXT , tagline TEXT, storyline TEXT, rating TEXT, genres TEXT, review_1 TEXT, review_2 TEXT, review_3 TEXT, review_4 TEXT, review_5 TEXT, status TEXT)")

    except:
        print("Table already exist")


def read_from_excel():
    excel.open_workbook("movies.xlsx")
    movies_table = excel.read_worksheet_as_table(header=True)
    excel.close_workbook
    return movies_table


def go_to_home_page():
    browser.click_element('xpath://*[@id="home_img_holder"]')


def type_and_search_movies(movie, cur, con, mid):
    movie_name = movie["Movie"]
    browser.input_text('xpath://*[@id="suggestion-search"]', movie_name)
    browser.click_element('suggestion-search-button')
    filter = "Movies"
    get_all = browser.get_text(
        'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')
    if get_all == filter:
        browser.click_element(
            'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')

    exact_movie_index = search_for_exact_match(movie_name)

    search_for_latest_movie_and_click_the_movie(
        exact_movie_index, movie_name, cur, con, mid)


def search_for_latest_movie_and_click_the_movie(exact_movie_index, movie_name, cur, con, mid):
    latest_release_year = 0
    latest_release_year_index = 0
    j = 0
    for j in exact_movie_index:
        j = str(j)

        movie_year = browser.get_text(
            'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + j + ']/div[2]/div/ul[1]/li')
        movie_year = int(movie_year)
        if movie_year > latest_release_year:
            latest_release_year = movie_year
            latest_release_year_index = j
    latest_release_year_index = str(latest_release_year_index)
    try:

        browser.click_element(
            'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + latest_release_year_index + ']/div[2]/div/a')
        extract_data(movie_name, cur, con, mid)
    except:
        status = "No exact match found"
        no_found(movie_name, status, cur, con, mid)


def extract_data(movie_name, cur, con, mid):
    status = "Success"
    try:
        rating = browser.get_text(ratingpath)
    except:
        rating = "--N/A--"

    try:
        storyline = browser.get_text(storylinepath)
    except:
        storyline = "--N/A--"

    browser.execute_javascript("window.scrollTo(0,3500)")
    time.sleep(2)

    try:
        tagline = browser.get_text(taglinepath1)
    except:
        try:
            tagline = browser.get_text(taglinepath2)
        except:
            tagline = "--N/A--"

    browser.execute_javascript("window.scrollTo(3500,0)")
    time.sleep(2)

    try:
        genres = browser.get_text(genrespath)
    except:
        genres = "--N/A--"

    browser.click_element(user_review_path)

    try:
        review1 = browser.get_text(review1path)
    except:
        review1 = "--N/A--"

    try:
        review2 = browser.get_text(review2path)
    except:
        review2 = "--N/A--"

    try:
        review3 = browser.get_text(review3path)
    except:
        review3 = "--N/A--"

    try:
        review4 = browser.get_text(review4path)
    except:
        review4 = "--N/A--"

    try:
        review5 = browser.get_text(review5path)
    except:
        review5 = "--N/A--"

    storyline = remove_punctuation(storyline)
    tagline = remove_punctuation(tagline)
    review1 = remove_punctuation(review1)
    review2 = remove_punctuation(review2)
    review3 = remove_punctuation(review3)
    review4 = remove_punctuation(review4)
    review5 = remove_punctuation(review5)

    # insert_into_database(mid, movie_name, cur, rating, storyline, tagline,
    #                      genres, review1, review2, review3, review4, review5, status, con)

    print(rating, "/n")
    print(storyline, "/n")
    print(tagline, "/n")
    print(genres, "/n")
    print(review1, "/n")
    print(review2, "/n")
    print(review3, "/n")
    print(review4, "/n")
    print(review5, "/n")


def remove_punctuation(string):
    import re
    pattern = r'[\"\',]'
    return re.sub(pattern, '', string)


def insert_into_database(mid, movie_name, cur, rating, storyline, tagline, genres, review1, review2, review3, review4, review5, status, con):
    cur.execute("INSERT INTO movie(mid,movie_name,tagline,storyline,rating,genres,review1,review2,review3,review4,review5,status) values"), (
        'mid', 'movie_name', 'tagline', 'storyline', 'rating', 'genres', 'review1', 'review2', 'review3', 'review4', 'review5', 'status')
    con.commit()


def no_found(movie_name, status, cur, con, mid):
    rating = "--N/A--"
    storyline = "--N/A--"
    tagline = "--N/A--"
    genres = "--N/A--"
    review1 = "--N/A--"
    review2 = "--N/A--"
    review3 = "--N/A--"
    review4 = "--N/A--"
    review5 = "--N/A--"

    # insert_into_database(mid, movie_name, cur, rating, storyline, tagline,
    #                      genres, review1, review2, review3, review4, review5, status, con)


def search_for_exact_match(movie_name):
    search_result_count = browser.get_element_count(
        'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li')

    if search_result_count > 5:
        search_result_count = 5
    search_result_count = search_result_count + 1

    exact_movie_index = []

    for i in range(1, search_result_count):

        i = str(i)
        movie_title = browser.get_text(
            'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + i + ']/div[2]/div/a')
        movie_title = str(movie_title)
        movie_title = movie_title.lower()
        movie_name = str(movie_name)
        movie_name = movie_name.lower()

        if movie_title == movie_name:
            exact_movie_index.append(i)
    return exact_movie_index


def main():
    open_the_browser()
    con = connect_database()
    cur = create_cursor(con)
    create_table(cur)
    try:
        movies_table = read_from_excel()
        mid = 1
        for movie in movies_table:
            type_and_search_movies(movie, cur, con, mid)
            mid = mid + 1
            go_to_home_page()

    finally:
        print("finally")


if __name__ == "__main__":
    main()
