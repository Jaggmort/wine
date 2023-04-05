import datetime
import pandas
import collections
import warnings
import os
from dotenv import load_dotenv
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def create_age_string():
    company_creation_date = 1920
    age = datetime.datetime.now().year - company_creation_date
    cut_age = age % 100
    if cut_age == 1:
        result = f'Уже {age} год с вами'
    if cut_age > 1 and cut_age < 5:
        result = f'Уже {age} года с вами'
    if cut_age > 4 or cut_age == 0:
        result = f'Уже {age} лет с вами'
    return result


def load_excel(data_file):
    xlsx = pandas.read_excel(data_file)
    xlsx = xlsx.fillna('')
    column_names = xlsx.columns.ravel()
    wines = collections.defaultdict(list)
    for wine_number in range(len(xlsx)):
        result = {'title': xlsx[column_names[1]][wine_number],
                  'sort': xlsx[column_names[2]][wine_number],
                  'price': xlsx[column_names[3]][wine_number],
                  'image': f'images/{xlsx[column_names[4]][wine_number]}',
                  'discont': xlsx[column_names[5]][wine_number],
                  }
        wines[xlsx[column_names[0]][wine_number]].append(result)
    return wines


def main():
    load_dotenv()
    warnings.filterwarnings("ignore")
    data_file = os.environ.get('XLSX_DATA_FILE')
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    wines = load_excel(data_file)
    rendered_page = template.render(
        age=create_age_string(),
        wines=wines,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
