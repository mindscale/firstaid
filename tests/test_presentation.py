from itertools import combinations
from pathlib import Path
import unittest

import pandas as pd
from firstaid.pptx import PPTX, LineChart, BarChart


def test_bulletpoint1():
    p = PPTX()
    p.add('This is a test', ['first bullet', 'second bullet'])
    p.save('test_bullet1.pptx')


def test_bulletpoint2():
    p = PPTX()
    p.add('This is a test', ['first bullet', 'second bullet'], ['third bullet', 'fourth bullet'])
    p.save('test_bullet2.pptx')


def test_image1(icecream_png):
    p = PPTX()
    p.add('This is a test', icecream_png)
    p.save('test_image1.pptx')


def test_image2(icecream_png):
    p = PPTX()
    p.add('This is a test', icecream_png, icecream_png)
    p.save('test_image2.pptx')


def test_table1():
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    p = PPTX()
    p.add('This is a test', df)
    p.save('test_table1.pptx')


def test_table2():
    df1 = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    df2 = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    p = PPTX()
    p.add('This is a test', df1, df2)
    p.save('test_table2.pptx')


def test_plot1():
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])

    p = PPTX()
    p.add('This is a test', LineChart(df))
    p.save('test_plot1.pptx')


def test_plot2():
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])

    p = PPTX()
    p.add('This is a test', LineChart(df), BarChart(df))
    p.save('test_plot2.pptx')


def test_mixed_combinations(icecream_png):
    bullet = ['first bullet', 'second bullet']
    img = icecream_png
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])

    all_cases = [bullet, img, df, BarChart(df)]
    comb = list(combinations(all_cases, 2))

    p = PPTX()
    for i, (left, right) in enumerate(comb):
        p.add('This is a test', left, right)
    p.save('test_combination.pptx')


class TestError(unittest.TestCase):
    def test_identify_tuple(self):
        """처음보는 자료형"""
        p = PPTX()
        with self.assertRaises(TypeError):
            p.identify_content(tuple([1, 2]))

    def test_bullet_list_elements(self):
        """['가나다', '마바사] 형태인지 확인"""
        p = PPTX()
        with self.assertRaises(TypeError):
            p.identify_content(['abc', 1, ['a', 'b', 'c']])

    def test_is_image_file(self):
        """'image.png' 형태인지 확인"""
        p = PPTX()
        with self.assertRaises(FileNotFoundError):
            p.identify_content('this is a string')


def test_getting_started_pics(bird_png, dog_png, horse_png, icecream_png):
    my_file = PPTX()
    my_file.add(title='제목입니다',
                content1=['첫 문장입니다', '두번째 문장입니다'])
    my_file.add(title='제목입니다',
                content1=['첫 문장입니다', '두번째 문장입니다'],
                content2=['first sentence', 'second sentence'])
    my_file.add('제목입니다',
                ['첫 문장입니다', '두번째 문장입니다'],
                ['first sentence', 'second sentence'])
    points = ['1. 첫 문장 입니다',
              '2. 두번째 문장입니다',
              '3. 세번째 문장입니다']
    my_file.add('문자열 슬라이드입니다',
                points)
    my_file.add('이미지 슬라이드입니다',
                icecream_png)
    df = pd.DataFrame([[10, 30, 50], [24, 13, 70]],
                      columns=['A사', 'B사', 'C사'],
                      index=['남', '여'])
    my_file.add('표 슬라이드 입니다', df)
    my_file.add('차트 슬라이드 입니다',
                BarChart(df))

    my_file.add('내가 좋아하는 디저트',
                ['아이스크림', '초콜릿'],  # 문자열
                icecream_png)  # 이미지
    my_file.add('동물들의 발자국',
                ['동물들의 발자국을 공부해봅시다'])
    # 두번째 슬라이드
    my_file.add('개', dog_png)
    # 세번째 슬라이드
    my_file.add('말', horse_png)
    # 네번째 슬라이드
    my_file.add('새', bird_png)
    my_file.save('example.pptx')
