from itertools import combinations

import pandas as pd
import pytest
from firstaid.pptx import PPTX


def test_bulletpoint1():
    p = PPTX()
    p.add('This is a test', ['first bullet', 'second bullet'])
    p.save('tests/presentation_test_results/test_bullet1.pptx')


def test_bulletpoint2():
    p = PPTX()
    p.add('This is a test', ['first bullet', 'second bullet'], ['third bullet', 'fourth bullet'])
    p.save('tests/presentation_test_results/test_bullet2.pptx')


def test_image1():
    p = PPTX()
    p.add('This is a test', 'tests/icecream.png')
    p.save('tests/presentation_test_results/test_image1.pptx')


def test_image2():
    p = PPTX()
    p.add('This is a test', 'tests/icecream.png', 'tests/icecream.png')
    p.save('tests/presentation_test_results/test_image2.pptx')


def test_table1():
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    p = PPTX()
    p.add('This is a test', df)
    p.save('tests/presentation_test_results/test_table1.pptx')


def test_table2():
    df1 = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    df2 = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    p = PPTX()
    p.add('This is a test', df1, df2)
    p.save('tests/presentation_test_results/test_table2.pptx')


def test_plot1():
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    d = {'df': df, 'chart_type': 'line'}

    p = PPTX()
    p.add('This is a test', d)
    p.save('tests/presentation_test_results/test_plot1.pptx')


def test_plot2():
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    d1 = {'df': df, 'chart_type': 'line'}
    d2 = {'df': df, 'chart_type': 'bar'}

    p = PPTX()
    p.add('This is a test', d1, d2)
    p.save('tests/presentation_test_results/test_plot2.pptx')


def test_mixed_combinations():
    bullet = ['first bullet', 'second bullet']
    img = 'tests/icecream.png'
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    d = {'df': df, 'chart_type': 'line'}

    all_cases = [bullet, img, df, d]
    comb = list(combinations(all_cases, 2))

    p = PPTX()
    for i, (left, right) in enumerate(comb):
        p.add('This is a test', left, right)
    p.save('tests/presentation_test_results/test_combination.pptx')


@pytest.mark.xfail(run=False)
def test_getting_started_pics():
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
                'tests/icecream.png')
    df = pd.DataFrame([[10, 30, 50], [24, 13, 70]],
                      columns=['A사', 'B사', 'C사'],
                      index=['남', '여'])
    my_file.add('표 슬라이드 입니다', df)
    chart_dict = {'df': df,  # pandas DataFrame 형식
                  'chart_type': 'line'}  # 'bar' 도 가능하다
    my_file.add('차트 슬라이드 입니다',
                chart_dict)

    my_file.add('내가 좋아하는 디저트',
                ['아이스크림', '초콜릿'],  # 문자열
                'tests/icecream.png')  # 이미지
    my_file.add('동물들의 발자국',
                ['동물들의 발자국을 공부해봅시다'])
    # 두번째 슬라이드
    my_file.add('개', 'tests/dog.png')
    # 세번째 슬라이드
    my_file.add('말', 'tests/horse.png')
    # 네번째 슬라이드
    my_file.add('새', 'tests/bird.png')
    my_file.save('tests/presentation_test_results/example.pptx')
