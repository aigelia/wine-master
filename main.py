from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas


def count_company_age():
    now = datetime.now().year
    company_found = 1920
    age = now - company_found
    return age


def generate_age_word(age):
    if 11 <= age % 100 <= 14:
        return "лет"
    elif age % 10 == 1:
        return "год"
    elif 2 <= age % 10 <= 4:
        return "года"
    else:
        return "лет"


def read_excel(file_name, sheet_name):
    excel_data = pandas.read_excel(file_name, sheet_name=sheet_name).fillna('')
    return excel_data.to_dict(orient="records")


def generate_wine_cards(wine_list):
    return [
        {
            "wine_name": wine["Название"],
            "grape_variety": wine["Сорт"],
            "wine_price": wine["Цена"],
            "image_name": wine["Картинка"],
            "name_category": wine["Категория"],
            "discount": wine["Акция"]
        }
        for wine in wine_list
    ]


def split_wine_categories(data):
    grouped = defaultdict(list)
    for item in data:
        category = item["name_category"]
        grouped[category].append(item)
    return dict(grouped)


def main():
    age = count_company_age()
    age_word = generate_age_word(age)

    wine_data = read_excel("wine_assortment.xlsx", "Лист1")
    wine_cards = generate_wine_cards(wine_data)
    categorized_wines = split_wine_categories(wine_cards)

    env = Environment(
        loader=FileSystemLoader("."),
        autoescape=select_autoescape(['html'])
    )
    template = env.get_template('template.html')

    rendered_page = template.render(
        company_age=age,
        age_word=age_word,
        categorized_wines=categorized_wines
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
