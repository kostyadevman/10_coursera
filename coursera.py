import requests
import bs4
import random
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


def get_html_from_url(coursera_url):
    return requests.get(coursera_url).content


def get_random_courses_references(coursera_html, course_count):
    courses_references = []
    coursera_soup = bs4.BeautifulSoup(
        coursera_html,
        'lxml'
    )
    coursera_loc_tags = coursera_soup.find_all('loc')
    for loc_tag in coursera_loc_tags:
        courses_references.append(loc_tag.text)
    return random.sample(
        courses_references,
        course_count
    )


def get_course_info(course_html):
    coursera_soup = bs4.BeautifulSoup(course_html, 'lxml')
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
        user_rating = None
    return {
        '1_start_date': start_date,
        '2_title': title,
        '3_language': language,
        '4_week_count': week_count,
        '5_user_rating': user_rating,
    }


def fill_xlsx(courses_info):
    header = [
        'Start date',
        'Course title',
        'Language',
        'Week count',
        'User rating (5 maximum)',
    ]
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'courses info'
    sheet.append(header)
    for course_info in courses_info:
        if not course_info['5_user_rating']:
            course_info['5_user_rating'] = 'No rating yet'
        sheet.append(
            [
                course_info['1_start_date'],
                course_info['2_title'],
                course_info['3_language'],
                course_info['4_week_count'],
                course_info['5_user_rating']
            ]
        )
    return workbook


if __name__ == '__main__':
    args = get_arguments()
    course_count = 20
    courses_info = []
    coursera_html = get_html_from_url(
        'https://www.coursera.org/sitemap~www~courses.xml'
    )
    courses_references = get_random_courses_references(
        coursera_html,
        course_count
    )

    for courses_reference in courses_references:
        course_html = get_html_from_url(courses_reference)
        course_info = get_course_info(course_html)
        courses_info.append(course_info)
    workbook = fill_xlsx(courses_info)
    workbook.save(args.output)
