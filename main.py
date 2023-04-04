import datetime
import pandas
import collections
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
    excel = pandas.read_excel('wine3.xlsx')
    excel = excel.fillna('None')
    column_names = excel.columns.ravel()
    default_dict = collections.defaultdict(list)
    for wine_number in range(len(excel)):
        result = {'title': excel[column_names[1]][wine_number],
                  'sort': excel[column_names[2]][wine_number],
                  'price': excel[column_names[3]][wine_number],
                  'image': f'images/{excel[column_names[4]][wine_number]}',
                  'discont': excel[column_names[5]][wine_number],
                  }
        default_dict[excel[column_names[0]][wine_number]].append(result)
    return default_dict

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
