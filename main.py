import argparse
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
from collections import defaultdict


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--filepath', type=str, default='wine.xlsx')

    return parser.parse_args()


def get_estimate_label():
    now = datetime.date.today().year
    estimate_year = 1920
    years_estimated = now - estimate_year

    if 4 < years_estimated % 100 < 21:
        return f'{years_estimated} лет'

    if years_estimated % 10 == 1:
        return f'{years_estimated} год'

    if 1 < years_estimated % 10 < 5:
        return f'{years_estimated} года'

    return f'{years_estimated} лет'


def read_excel(filepath):
    products = pandas.read_excel(
        filepath,
        na_values=' ',
        keep_default_na=False,
    ).to_dict(orient='records')
    products.sort(key=lambda i: i['Категория'])
    structured_products = defaultdict(list)

    for product in products:
        structured_products[product['Категория']].append(product)

    return structured_products


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    argument = parse_args()
    assortment = read_excel(argument.filepath)
    estimate_label = get_estimate_label()

    rendered_page = template.render(
        assortment=assortment,
        estimate_label=estimate_label,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
