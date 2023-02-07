import requests
from bs4 import BeautifulSoup
import re
import openpyxl as op

excel = op.Workbook()
sheet = excel.active
sheet.title = "Popular Movie List"
sheet.append(["index","movie_name","ratings","castings","oneline"])
try:

    req = requests.get("https://timesofindia.indiatimes.com/entertainment/tamil/kollywood/top-20-best-tamil-movies-of-2022")

    soup = BeautifulSoup(req.content, "html.parser")
    #print(soup.prettify())
    swap = soup.find("div", class_="wrapper clearfix").find_all("div", class_="topten_movie_block")
    for movie in swap:
        #print(movie)
        index = movie.find("div", class_="number_block").find("span").get_text(strip=True)
        movie_name = movie.find("div", class_="topten_movies_content").h2.get_text(strip=True)
        movie_name_1 = re.sub("\W", " ", movie_name)
       # image = movie.find("div", class_="topten_movie_img").img
        ratings = movie.find("span", class_="topten_ratemovie").span.text
        castings = movie.find("h3").text
        oneline = movie.find("h3").findNext("h3").text

        print(index,movie_name_1,ratings,castings,oneline)#,image)

        sheet.append([index,movie_name_1,ratings,castings,oneline])#,image])
except:
    print("error")

excel.save("moviies.xlsx")
