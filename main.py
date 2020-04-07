import datetime
import argparse
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_company_age():
    company_foundation_date = datetime.datetime(year=1920, month=1, day=1)
    current_date = datetime.datetime.now()
    company_age = current_date.year - company_foundation_date.year

    return company_age


def main():
    parser = argparse.ArgumentParser(
        description='This script parse information about wines and run website local by this link http://127.0.0.1:8000/.'
    )
    parser.add_argument('--wine_path', help='Input path to wine.xlsx file. Default: "wine.xlsx"')
    args = parser.parse_args()

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    company_age = get_company_age()

    path_to_wines = args.wine_path if args.wine_path else 'wine.xlsx'
    wines_df = pandas.read_excel(
        path_to_wines,
        na_values=None,
        keep_default_na=False
    )
    wines = wines_df.to_dict(orient='records')
    wines_categories = set(wines_df['Категория'].to_list())

    rendered_page = template.render(
        company_age=company_age,
        wines=wines,
        wines_categories = wines_categories
    )

    with open('index.html', 'w', encoding='utf8') as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
