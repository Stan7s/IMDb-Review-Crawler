# coding=utf-8
import requests
from urllib.parse import urljoin
from bs4 import BeautifulSoup
import xlwt
import re


base_url = "https://www.imdb.com/"
key = ""

# Change the movie review page url here
url = 'https://www.imdb.com/title/tt0120731/reviews?ref_=tt_ql_3'


f = xlwt.Workbook()
sheet1 = f.add_sheet('Movie Reviews', cell_overwrite_ok=True)
row = ["Title", "Author", "Date", "Up Vote", "Total Vote", "Rating", "Review"]
for i in range(0, len(row)):
    sheet1.write(0, i, row[i])

MAX_CNT = 100
cnt = 1

print("url = ", url)
res = requests.get(url)
res.encoding = 'utf-8'
soup = BeautifulSoup(res.text, "lxml")


for item in soup.select(".lister-list"):
    title = item.select(".title")[0].text
    author = item.select(".display-name-link")[0].text
    date = item.select(".review-date")[0].text
    votetext = item.select(".text-muted")[0].text
    upvote = re.findall(r"\d+",votetext)[0]
    totalvote = re.findall(r"\d+", votetext)[1]
    rating = item.select("span.rating-other-user-rating > span")
    if len(rating) == 2:
        rating = rating[0].text
    else:
        rating = ""
    review = item.select(".text")[0].text
    row = [title, author, date, upvote, totalvote, rating, review]
    for i in range(0, len(row)):
        sheet1.write(cnt, i, row[i])
    cnt = cnt + 1

load_more = soup.select(".load-more-data")
flag = True
if len(load_more):
    ajaxurl = load_more[0]['data-ajaxurl']
    base_url = base_url + ajaxurl + "?ref_=undefined&paginationKey="
    key = load_more[0]['data-key']
else:
    flag = False

while flag:
    url = base_url + key
    print("url = ", url)
    res = requests.get(url)
    res.encoding = 'utf-8'
    soup = BeautifulSoup(res.text, "lxml")
    for item in soup.select(".lister-item-content"):
        title = item.select(".title")[0].text
        author = item.select(".display-name-link")[0].text
        date = item.select(".review-date")[0].text
        votetext = item.select(".text-muted")[0].text
        vote = re.findall(r"\d+", votetext)[0]
        totalvote = re.findall(r"\d+", votetext)[1]
        rating = item.select("span.rating-other-user-rating > span")
        if len(rating) == 2:
            rating = rating[0].text
        else:
            rating = ""
        review = item.select(".text")[0].text
        row = [title, author, date, vote, totalvote, rating, review]
        for i in range(0, len(row)):
            sheet1.write(cnt, i, row[i])
        if cnt >= MAX_CNT:
            break
        cnt = cnt + 1
    if cnt >= MAX_CNT:
        break
    load_more = soup.select(".load-more-data")
    if len(load_more):
        key = load_more[0]['data-key']
    else:
        flag = False

f.save('Review.xls')
print(cnt, "reviews saved.")

