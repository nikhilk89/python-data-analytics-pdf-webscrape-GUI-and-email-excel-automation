
import csv
import re
import requests
from bs4 import BeautifulSoup

#https://books.toscrape.com/
urls = [
    "https://books.toscrape.com/catalogue/page-1.html",
    "https://books.toscrape.com/catalogue/page-2.html",
]

rating_map = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

#encoding='utf-8-sig' for excel to showpund sign
file = open("books.csv", mode="w", newline="", encoding="utf-8-sig")
writer = csv.writer(file)

writer.writerow(["bookname", "rating", "price", "in_stock"])

for url in urls:
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")

    articles = soup.find_all("article", class_="product_pod")

    for article in articles:
        # Book Name
        title_tag = article.h3.find("a")
        bookname = (
            title_tag["title"] if "title" in title_tag.attrs else title_tag.text
        )

        # Rating
        rating_classes = article.find("p", class_="star-rating")["class"]
        rating = rating_map.get(rating_classes[1], rating_classes[1])

        # price ccy change
        raw_price = article.find("p", class_="price_color").text
        clean_number = re.sub(r"[^\d.]", "", raw_price)
        price = f"£{clean_number}"

        # In Stock
        stock_text = article.find("p", class_="instock availability").text
        in_stock = "In stock" in stock_text

        writer.writerow([bookname, rating, price, in_stock])

file.close()

print("Done! Fixed currency saved to books.csv")