import json
import requests
from argparse import ArgumentParser
from lxml import etree
from random import sample
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_courses_file():
    response = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    if response.status_code == 200:
        return response.content


def get_courses_list(courses_xml, num_courses=20):
    root = etree.fromstring(courses_xml)
    courses_list = [leaf.find('{*}loc').text for leaf in root]
    courses_list = sample(courses_list, num_courses)
    return courses_list


def get_course_start_date(json_info):
    if not json_info:
        return None
    json_info = json.loads(json_info.text)
    course_instance = json_info.get('hasCourseInstance').pop()
    return course_instance.get('startDate')


def get_course_info(course_slug):
    response = requests.get(course_slug)
    soup = BeautifulSoup(response.content, 'html.parser')

    title = soup.find('div', class_='rc-CTANavItem').text
    language = soup.find('div', class_='language-info').text
    duration = soup.find('div', class_='rc-WeekView')
    duration = duration.find_all('div', class_='week') if duration else None
    avg_score = soup.find('div', class_='ratings-text bt3-visible-xs')
    json_info = soup.find('script', type='application/ld+json')
    upcoming_session = get_course_start_date(json_info)

    course_info = {
        'A': title if title else soup.title.text,
        'B': language,
        'C': upcoming_session,
        'D': '%d weeks' % len(duration) if duration else 'Non specified',
        'E': avg_score.text if avg_score else 'Not rated yet'
    }
    return course_info


def output_courses_info_to_xlsx(courses_info, filepath):
    xlsx_file = Workbook()
    worksheet = xlsx_file.active
    header = ['Title', 'Language', 'Upcoming Session', 'Duration', 'Average Score']
    worksheet.append(header)
    for course in courses_info:
        worksheet.append(course)
    xlsx_file.save(filepath)


if __name__ == '__main__':
    parser = ArgumentParser(description='Retrieve info from n specified Coursera courses')
    parser.add_argument('--n_courses', nargs='?', type=int, default=20)
    parser.add_argument('--output', nargs='?', type=str, default='courses_info.xlsx')
    args = parser.parse_args()

    courses_xml = get_courses_file()
    courses_list = get_courses_list(courses_xml, args.n_courses)
    courses_info = []
    for count, course in enumerate(courses_list):
        print('Proceeding course {0}/{1}...'.format(count + 1, len(courses_list)), end='\r')
        courses_info.append(get_course_info(course))
    output_courses_info_to_xlsx(courses_info, filepath=args.output)
    print('Results saved to "%s"' % args.output)
