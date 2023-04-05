import datetime
import pandas
import collections
import warnings
import os
from dotenv import load_dotenv
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def check_age():
    age = datetime.datetime.now().year - 1920
    cut_age = age % 100
    if cut_age == 1:
        result = f'Уже {age} год с вами'
    if cut_age > 1 and cut_age < 5:
        result = f'Уже {age} года с вами'
    if cut_age > 4 or cut_age == 0:
        result = f'Уже {age} лет с вами'
    return result


def load_excel():
    load_dotenv()
    data_file = os.environ.get('DATA_FILE')
    data_from_xlsx = pandas.read_excel(data_file)
    data_from_xlsx = data_from_xlsx.fillna('None')
    column_names = data_from_xlsx.columns.ravel()
    processed_wine_data = collections.defaultdict(list)
    for wine_number in range(len(data_from_xlsx)):
        result = {'title': data_from_xlsx[column_names[1]][wine_number],
                  'sort': data_from_xlsx[column_names[2]][wine_number],
                  'price': data_from_xlsx[column_names[3]][wine_number],
                  'image': f'images/{data_from_xlsx[column_names[4]][wine_number]}',
                  'discont': data_from_xlsx[column_names[5]][wine_number],
                  }
        processed_wine_data[data_from_xlsx[column_names[0]][wine_number]].append(result)
    return processed_wine_data

def main():
    warnings.filterwarnings("ignore")
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    wines = load_excel()
    rendered_page = template.render(
        age=f'{check_age()}',
        wines=wines,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

if __name__ == '__main__':
    main()