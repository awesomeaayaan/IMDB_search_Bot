from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import logging
import time
import re
import sqlite3

#logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(message)s')
class MovieScraper:
    logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(message)s')
    def __init__(self):
        self.browser_lib = Selenium(auto_close=False)
        self.FILE_NAME = "movies.xlsx"
        self.url = 'https://www.imdb.com/'
        self.excel = Files()
        
        self.ratingpath = 'xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[2]/div/div[1]/a/span/div/div[2]/div[1]/span[1]'
        self.storylinepath = 'xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/p'
        self.taglinepath1 = 'xpath://*[@id="__next"]/main/div/section[1]/div/section/div/div[1]/section[6]/div[2]/ul[2]/li[1]/div/div/ul'
        self.taglinepath2 = 'xpath://*[@id="__next"]/main/div/section[1]/div/section/div/div[1]/section[7]/div[2]/ul[2]/li[1]/div/div/ul'
        self.genrespath = 'xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/div[1]/div[2]'
        self.user_review_path = 'xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[1]/div/div[2]/ul/li[2]/a'
        self.review1path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[1]/div/div[1]/a'
        self.review2path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[2]/div[1]/div[1]/a'
        self.review3path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[3]/div[1]/div[1]/a'
        self.review4path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[4]/div[1]/div[1]/a'
        self.review5path = 'xpath://*[@id="main"]/section/div[2]/div[2]/div[5]/div[1]/div[1]/a'
        self.con = sqlite3.connect("moviee.db")
        self.cur = self.con.cursor()

    def open_browser(self):
        self.browser_lib.open_available_browser(self.url, maximized=True)

    def create_table(self):
        try:
            self.cur.execute(
                "CREATE TABLE movieee(Mid INTEGER PRIMARY KEY AUTOINCREMENT, movie_name TEXT , tagline TEXT, storyline TEXT, rating TEXT, genres TEXT, review_1 TEXT, review_2 TEXT, review_3 TEXT, review_4 TEXT, review_5 TEXT, status TEXT)")
        except:
            print("Table already exists")

    def read_from_excel(self):
        self.excel.open_workbook(self.FILE_NAME)
        movies_table = self.excel.read_worksheet_as_table(header=True)
        self.excel

    def insert_movies(self, movie):
        movie_name = movie['Movie']
        self.browser_lib.input_text('//*[@id="suggestion-search"]', movie_name)
        self.browser_lib.click_button('suggestion-search-button')
        filter_movie = 'Movies'
        get_all = self.browser_lib.get_text('//*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')

        if get_all == filter_movie:
            self.browser_lib.click_element('//*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')
        exact_movie_index = self.search_for_exact_match(movie_name)
        self.search_for_latest_movie_and_click_the_movie(exact_movie_index, movie_name)

    def extract_data(self, movie_name):
        status = "Success"
        try:
            rating = self.browser_lib.get_text(self.ratingpath)
        except:
            rating = 'No rating'
        try:
            storyline = self.browser_lib.get_text(self.storylinepath)
        except:
            storyline = 'no storyline'
        try:
            genres = self.browser_lib.get_text(self.genrespath)
        except:
            genres = 'no genres'
        
        self.browser_lib.execute_javascript("window.scrollTo(0,3500)")
        time.sleep(2)
        try:
            tagline = self.browser_lib.get_text(self.taglinepath1)
        except:
            try:
                tagline = self.browser_lib.get_text(self.taglinepath2)
            except:
                tagline = "--N/A--"
        self.browser_lib.execute_javascript("window.scrollTo(3500,0)")
        time.sleep(2)
        #print(tagline)
        #tagline = browser_lib.get_text(taglinepath1)
        self.browser_lib.click_element(self.user_review_path)
        try:
            review1 = self.browser_lib.get_text(self.review1path)
            rev1 = self.remove_punctuations(review1)
            print(rev1)
        except:
            print('review not find')
        

        try:
            review2 = self.browser_lib.get_text(self.review2path)
            rev2 = self.remove_punctuations(review2)
            print(rev2)
        except:
            print('review not find')

        try:
            review3 = self.browser_lib.get_text(self.review3path)
            rev3 = self.remove_punctuations(review3)
            print(rev3)
        except:
            print('review not found')
        try:
            review4 = self.browser_lib.get_text(self.review4path)
            rev4 = self.remove_punctuations(review4)
            print(rev4)
        except:
            print('review not found')
        try:
            review5 = self.browser_lib.get_text(self.review5path)
            rev5 = self.remove_punctuations(review5)
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
        self.insert_into_table(movie_data)

    def search_for_latest_movie_and_click_the_movie(self,exact_movie_index, movie_name):
        latest_release_year = 0
        latest_release_year_index = 0
        j = 0
        for j in exact_movie_index:
            j = str(j)

            movie_year = self.browser_lib.get_text(
                'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + j + ']/div[2]/div/ul[1]/li')
            movie_year = int(movie_year)
            if movie_year > latest_release_year:
                latest_release_year = movie_year
                latest_release_year_index = j
        latest_release_year_index = str(latest_release_year_index)
        try:

            self.browser_lib.click_element(
                'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + latest_release_year_index + ']/div[2]/div/a')
            self.extract_data(movie_name)
        except:
            status = "No exact match found"
            self.no_movie_found(movie_name,status)

    def no_movie_found(self,movie_name, status):
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
        self.insert_into_table(movie_data)
        
    def search_for_exact_match(self,movie_name):
        search_result_count = self.browser_lib.get_element_count(
            'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li')

        if search_result_count > 5:
            search_result_count = 5
        search_result_count = search_result_count + 1

        exact_movie_index = []

        for i in range(1, search_result_count):

            i = str(i)
            movie_title = self.browser_lib.get_text(
                'xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[' + i + ']/div[2]/div/a')
            movie_title = str(movie_title)
            movie_title = movie_title.lower()
            movie_title = self.remove_punctuations(movie_title)
            movie_title = movie_title.strip()

            movie_name = str(movie_name)
            movie_name = movie_name.lower()
            movie_name = self.remove_punctuations(movie_name)
            movie_name = movie_name.strip()

            if movie_title == movie_name:
                exact_movie_index.append(i)
            
        return exact_movie_index

    def insert_into_table(self,movie_data):
        insert_sql = """INSERT INTO movieee(movie_name, 
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
        self.cur.execute(insert_sql, (movie_data["movie_name"],
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
        self.con.commit()


    def remove_punctuations(string):
    
        pattern = r'[\"\',]'
        return re.sub(pattern, '', string)


if __name__ == "__main__":
    movie_scraper = MovieScraper()
    movie_scraper.open_browser()
    movie_scraper.create_table()
    movies_table = movie_scraper.read_from_excel()
    for movie in movies_table:
        movie_scraper.insert_movies(movie)