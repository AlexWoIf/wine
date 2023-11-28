from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas
import collections


FOUNDATION_YEAR = 1920


def year_with_tail(num):
    tail = 'год'
    if (((num % 10) == 0)
            or ((num % 10) in range(5, 10))
            or ((num % 100) in range(11, 15))):
        tail = 'лет'
    elif (num % 10) in range(2, 5):
        tail = 'года'
    return (f"{num} {tail}")


def load_wines_from_xlsx(filename):
    data = (
        pandas.read_excel(filename, sheet_name='Лист1')
        .fillna('').to_dict(orient='records')
    )
    result = collections.defaultdict(list)
    for record in data:
        key = record.pop('Категория')
        result[key].append(record)
    return result


if __name__ == '__main__':
    env = Environment(loader=FileSystemLoader('.'),
                      autoescape=select_autoescape(['html', 'xml']))

    template = env.get_template('template.html')

    rendered_page = template.render(
        years=year_with_tail(datetime.now().year - FOUNDATION_YEAR),
        wines=load_wines_from_xlsx('wine3.xlsx'),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
