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
