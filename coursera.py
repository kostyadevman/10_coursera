import requests
import bs4
from random import randint
import openpyxl
import argparse


def get_arguments():
    parser = argparse.ArgumentParser('Get info from coursera.org')
    parser.add_argument(
        '-o',
        '--output',
        type=str,
        required=True,
        help='path to output .xlsx file'
    )
    return parser.parse_args()


def get_courses_list(course_count):
    courses = []
    coursera_response = requests.get(
        'https://www.coursera.org/sitemap~www~courses.xml'
    )
    coursera_soup = bs4.BeautifulSoup(
        coursera_response.content,
        'lxml'
    )
    coursera_loc_tags = coursera_soup.find_all('loc')
    for loc_tag in coursera_loc_tags:
        courses.append(loc_tag.text)
    for course_count in range(0, course_count):
        yield courses[randint(0, len(courses))]


def get_course_info(course_url):
    course_response = requests.get(course_url)
    coursera_soup = bs4.BeautifulSoup(course_response.content, 'lxml')
    title = coursera_soup.find(
        'h1',
        class_='title display-3-text'
    ).text
    start_date = coursera_soup.find(
        'div',
        class_='startdate rc-StartDateString caption-text'
    ).text.replace('Started', '').replace('Starts', '')
    language = coursera_soup.find(
        'div',
        class_='rc-Language'
    ).contents[1]
    week_count = len(coursera_soup.find_all('div', class_='week'))
    user_rating_tag = coursera_soup.find(
        'div',
        class_='ratings-text bt3-visible-xs'
    )
    if user_rating_tag:
        user_rating = user_rating_tag.contents[0].text.replace(
            'stars',
            ''
        )
    else:
        user_rating = '-'
    return [start_date, title, language, week_count, user_rating]


def output_courses_info_to_xlsx(filepath, courses_info):
    header = ['Start date',
               'Course title',
               'Language',
               'Week count',
               'User rating (5 maximum)']
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'courses info'
    sheet.append(header)
    for course_info in courses_info:
        sheet.append(course_info)
    workbook.save(filepath)


if __name__ == '__main__':
    args = get_arguments()
    course_count = 20
    courses_info = []
    courses = get_courses_list(course_count)
    for course in courses:
        courses_info.append(get_course_info(course))
    output_courses_info_to_xlsx(args.output, courses_info)
