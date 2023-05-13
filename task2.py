# from RPA.Browser.Selenium import Selenium
# from RPA.Excel.Files import Files
# import logging
# browser_lib = Selenium(auto_close=False)
# FILE_NAME = "movies.xlsx"

# url = 'https://www.imdb.com/'
# excel = Files()
# logging.basicConfig(level=logging.INFO,format='%(levelname)s:%(message)s')


# def open_browser():
#     # use headless=True so robot work without opening browser, removi maximize argument
#     # maximize help robot to open browser in full screen
#     browser_lib.open_available_browser(url, maximized=True)

# def read_from_excel():
#     excel.open_workbook('movies.xlsx')
#     movies_table = excel.read_worksheet_as_table(header=True)
#     excel.close_workbook
#     return movies_table

# def insert_movies(movie):
#     movie_name = movie['Movie']
#     browser_lib.input_text('//*[@id="suggestion-search"]',movie_name)
#     browser_lib.click_button('suggestion-search-button')
#     filter_movie = 'Movies'
#     get_all = browser_lib.get_text('//*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')

#     if filter_movie == 'Movies':
#         browser_lib.click_element('//*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]')
    
#     # exact_moves = search_exact_movie(movie_name)
#     # search_latest_movie_and_click(exact_moves,movie_name)