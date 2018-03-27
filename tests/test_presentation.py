from itertools import combinations

import pandas as pd

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
    p.add('This is a test', 'tests/image.png')
    p.save('tests/presentation_test_results/test_image1.pptx')


def test_image2():
    p = PPTX()
    p.add('This is a test', 'tests/image.png', 'tests/image.png')
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
    img = 'tests/image.png'
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c'], index=['남', '여'])
    d = {'df': df, 'chart_type': 'line'}

    all_cases = [bullet, img, df, d]
    comb = list(combinations(all_cases, 2))

    p = PPTX()
    for i, (left, right) in enumerate(comb):
        p.add('This is a test', left, right)
    p.save('tests/presentation_test_results/test_combination.pptx')
