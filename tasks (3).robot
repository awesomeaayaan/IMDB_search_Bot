# *** Settings ***
# Documentation       Search movies in the provided excel
# ...                 Search for exact matches only
# ...                 Extract rating, storyline, tagline, genres and top 5 reviews
# ...                 Save them to sqlite database (id, movie_name, ratings, storyline,
# ...                 tagline, genres, review_1, review_2, review_3, review_4, review_5, status)
# ...                 Status : No exact match found if not found, otherwise success in status

# Library             RPA.Browser.Selenium
# Library             RPA.Excel.Files
# Library             RPA.RobotLogListener
# Library             Collections
# Library             RPA.FileSystem
# Library             String
# Library             RPA.Database
# Library             DatabaseLibrary


# *** Variables ***
# ${rating_path}
# ...                     xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[2]/div/div[1]/a/span/div/div[2]/div[1]/span[1]
# ${storyline_path}
# ...                     xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/p
# ${tagline_path}
# ...                     xpath://*[@id="__next"]/main/div/section[1]/div/section/div/div[1]/section[6]/div[2]/ul[2]/li[1]/div/div/ul
# ${tagline_path2}
# ...                     xpath://*[@id="__next"]/main/div/section[1]/div/section/div/div[1]/section[7]/div[2]/ul[2]/li[1]/div/div/ul
# ${genres_path}
# ...                     xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/div[1]/div[2]
# ${user_review_path}
# ...                     xpath://*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[1]/div/div[2]/ul/li[2]/a
# ${review1_path}
# ...                     xpath://*[@id="main"]/section/div[2]/div[2]/div[1]/div/div[1]/a
# ${review2_path}
# ...                     xpath://*[@id="main"]/section/div[2]/div[2]/div[2]/div[1]/div[1]/a
# ${review3_path}
# ...                     xpath://*[@id="main"]/section/div[2]/div[2]/div[3]/div[1]/div[1]/a
# ${review4_path}
# ...                     xpath://*[@id="main"]/section/div[2]/div[2]/div[4]/div[1]/div[1]/a
# ${review5_path}
# ...                     xpath://*[@id="main"]/section/div[2]/div[2]/div[5]/div[1]/div[1]/a
# ${dbname}               movie
# ${tablename}            movies


# *** Tasks ***
# Establish Connection with database
#     Establish Connection
#     TRY
#         Create a table
#     EXCEPT
#         Log To Console    Table already exist
#     END

# Search movies in provided excel
#     Open the website
#     Open Excel file and read movie name and store in database


# *** Keywords ***
# Establish Connection
#     Connect To Database Using Custom Params
#     ...    sqlite3
#     ...    database="./${dbname}.db", isolation_level=None

# Create a table
#     ${create_table_sql}    Catenate    SEPARATOR=    \n
#     ...    CREATE TABLE ${tablename}(
#     ...    id INTEGER PRIMARY KEY AUTOINCREMENT,
#     ...    movie_name TEXT ,
#     ...    tagline TEXT,
#     ...    storyline TEXT,
#     ...    rating TEXT,
#     ...    genres TEXT,
#     ...    review_1 TEXT,
#     ...    review_2 TEXT,
#     ...    review_3 TEXT,
#     ...    review_4 TEXT,
#     ...    review_5 TEXT,
#     ...    status TEXT
#     ...    );
#     Execute Sql String    ${create_table_sql}

# Open the website
#     Open Available Browser    https://www.imdb.com/

# Open Excel file and read movie name and store in database
#     Open Workbook    movies.xlsx
#     ${movie_table}    Read Worksheet As Table    header=True
#     Close Workbook
#     ${id}    Set Variable    1

#     FOR    ${movie}    IN    @{movie_table}
#         Type and search movies    ${movie}    ${id}
#         Go to homepage
#     END

# Go to homepage
#     Click Element    xpath://*[@id="home_img_holder"]

# Type and search movies
#     [Arguments]    ${movie}    ${id}
#     Input Text    xpath://*[@id="suggestion-search"]    ${movie}[Movie]
#     Click Button    suggestion-search-button
#     ${check_movies}    Set Variable    Movies
#     ${check_movies}    Convert To String    ${check_movies}
#     ${all}    Get Text
#     ...    xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]
#     IF    "${all}" == "${check_movies}"
#         Click Element
#         ...    xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[2]/div[2]/section[1]/div[2]/div[2]/a[1]
#     END

#     ${movie_name}    Set Variable    ${movie}[Movie]

#     TRY
#         ${Exact_movie_Index}    Search for exact match    ${movie}
#         Search for latest movie and click the movie    ${Exact_movie_Index}    ${movie}
#     EXCEPT    message
#         Log To Console    Couldnt find
#     END

# Search for exact match
#     [Arguments]    ${movie}
#     ${search_result_count}    Get Element Count
#     ...    xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li
#     IF    ${search_result_count} > 5
#         ${search_result_count}    Set Variable    5
#     END
#     ${search_result_count}    Evaluate    ${search_result_count} + 1

#     ${Exact_movie_Index}    Create List
#     FOR    ${i}    IN RANGE    1    ${search_result_count}
#         ${movie_title}    Get Text
#         ...    xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[${i}]/div[2]/div/a
#         ${movie_title}    Convert To Lower Case    ${movie_title}
#         ${movie_title}    Convert To String    ${movie_title}
#         ${lmovie}    Convert To Lower Case    ${movie}[Movie]
#         ${lmovie}    Convert To String    ${lmovie}

#         IF    "${movie_title}" == "${lmovie}"
#             Append To List    ${Exact_movie_Index}    ${i}
#         END
#     END
#     RETURN    ${Exact_movie_Index}

# Search for latest movie and click the movie
#     [Arguments]    ${Exact_movie_Index}    ${movie}
#     ${latest_release_year}    Set Variable    0
#     ${latest_release_year_Index}    Set Variable    0
#     ${j}    Set Variable    0
#     FOR    ${j}    IN    @{Exact_movie_Index}
#         ${movie_year}    Get Text
#         ...    //*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[${j}]/div[2]/div/ul[1]/li
#         IF    ${latest_release_year} < ${movie_year}
#             ${latest_release_year}    Set Variable    ${movie_year}
#             ${latest_release_year_Index}    Set Variable    ${j}
#         END
#     END
#     TRY
#         Click Element
#         ...    xpath://*[@id="__next"]/main/div[2]/div[3]/section/div/div[1]/section[2]/div[2]/ul/li[${latest_release_year_Index}]/div[2]/div/a
#         Extract data    ${movie}
#     EXCEPT
#         ${status}    Set Variable    No exact match found
#         Except variable    ${movie}    ${status}
#     END

# Extract data
#     [Arguments]    ${movie}
#     ${status}    Set Variable    Success
#     ${validate_rating}    Get WebElement    ${rating_path}
#     ${validate_storyline}    Get WebElement    ${storyline_path}

#     ${validate_genres}    Get WebElement    ${genres_path}

#     IF    ${validate_rating}
#         ${rating}    Get Text    ${rating_path}
#     ELSE
#         ${rating}    Set Variable    ---
#     END

#     IF    ${validate_storyline}
#         ${storyline}    Get Text    ${storyline_path}
#     ELSE
#         ${storyline}    Set Variable    ---
#     END

#     # FOR    ${a}    IN RANGE    1    100
#     #    Press Keys    NONE    ARROW_DOWN
#     # END

#     Execute Javascript    window.scrollTo(0,3500)
#     Sleep    2
#     ${validate_tagline}    Get WebElements    ${tagline_path}
#     ${tagline}    Set Variable    NONE
#     IF    ${validate_tagline}
#         ${tagline}    Get Text    ${tagline_path}
#     ELSE
#         TRY
#             ${tagline}    Get Text    ${tagline_path2}
#         EXCEPT    message
#             ${tagline}    Set Variable    ---
#         END
#     END

#     Execute Javascript    window.scrollTo(3500,0)
#     Sleep    2

#     IF    ${validate_genres}
#         ${genres}    Get Text    ${genres_path}
#     ELSE
#         ${genres}    Set Variable    ---
#     END

#     Click Element    ${user_review_path}
#     ${validate_review1}    Get WebElement    ${review1_path}
#     ${validate_review2}    Get WebElement    ${review2_path}
#     ${validate_review3}    Get WebElement    ${review3_path}
#     ${validate_review4}    Get WebElement    ${review4_path}
#     ${validate_review5}    Get WebElement    ${review5_path}
#     IF    ${validate_review1}
#         ${review1}    Get Text    ${review1_path}
#     ELSE
#         ${review1}    Set Variable    ---
#     END

#     IF    ${validate_review1}
#         ${review2}    Get Text    ${review2_path}
#     ELSE
#         ${review2}    Set Variable    ---
#     END

#     IF    ${validate_review1}
#         ${review3}    Get Text    ${review3_path}
#     ELSE
#         ${review3}    Set Variable    ---
#     END

#     IF    ${validate_review1}
#         ${review4}    Get Text    ${review4_path}
#     ELSE
#         ${review4}    Set Variable    ---
#     END

#     IF    ${validate_review1}
#         ${review5}    Get Text    ${review5_path}
#     ELSE
#         ${review5}    Set Variable    ---
#     END

#     ${storyline}    Remove Punctuations    ${storyline}
#     ${tagline}    Remove Punctuations    ${tagline}
#     ${review1}    Remove Punctuations    ${review1}
#     ${review2}    Remove Punctuations    ${review2}
#     ${review3}    Remove Punctuations    ${review3}
#     ${review4}    Remove Punctuations    ${review4}
#     ${review5}    Remove Punctuations    ${review5}
#     ${sqlstring}    Set Variable
#     ...    INSERT INTO ${tablename} (movie_name, tagline, storyline, rating, genres, review_1, review_2, review_3, review_4, review_5, status) VALUES ("${movie}[Movie]", "${tagline}", '${storyline}', "${rating}", "${genres}", "${review_1}", "${review_2}", "${review_3}", "${review_4}", "${review_5}", "${status}")
#     Execute Sql String    ${sqlstring}

# Except variable
#     [Arguments]    ${movie}    ${status}
#     ${rating}    Set Variable    ---
#     ${storyline}    Set Variable    ---
#     ${tagline}    Set Variable    ---
#     ${genres}    Set Variable    ---
#     ${review1}    Set Variable    ---
#     ${review2}    Set Variable    ---
#     ${review3}    Set Variable    ---
#     ${review4}    Set Variable    ---
#     ${review5}    Set Variable    ---

#     ${sqlstring}    Set Variable
#     ...    INSERT INTO ${tablename} (movie_name, tagline, storyline, rating, genres, review_1, review_2, review_3, review_4, review_5, status) VALUES ("${movie}[Movie]", "${tagline}", "${storyline}", "${rating}", "${genres}", "${review_1}", "${review_2}", "${review_3}", "${review_4}", "${review_5}", "${status}")
#     Execute Sql String    ${sqlstring}

# Remove Punctuations
#     [Arguments]    ${string}
#     ${pattern}    Set Variable    [\"']
#     ${result}    Replace String Using Regexp    ${string}    ${pattern}    ${EMPTY}
#     RETURN    ${result}
