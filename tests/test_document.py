import pandas as pd

from firstaid.docx import DOCX


def test_doc_title():
    doc = DOCX()
    doc.add('1. 제목입니다')
    doc.blank_line()
    doc.add('2. 제목입니다')
    doc.save('test_title.docx')


def test_doc_paragraph():
    doc = DOCX()
    doc.add(['첫번째 문단입니다. ' * 30])
    doc.add(['두번째 문단입니다. ' * 45])
    doc.page_break()
    doc.add(['첫번째 문단입니다. ' * 30])
    doc.add(['두번째 문단입니다. ' * 45])
    doc.save('test_paragraph.docx')


def test_doc_image(icecream_png):
    doc = DOCX()
    doc.add(image=icecream_png)
    doc.add(['아이스크림입니다'])
    doc.add(image=icecream_png)
    doc.save('test_image.docx')


def test_doc_table():
    doc = DOCX()
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]],
                      columns=['A사', 'B사', 'C사'],
                      index=['남', '여'])
    doc.add(df)
    doc.add(['표입니다'])
    doc.add(df)
    doc.save('test_table.docx')


def test_doc_getting_started_pages1(icecream_png):
    my_file = DOCX()  # 작업 시작

    my_file.add(content=['첫 문장입니다', '두번째 문장입니다'])
    my_file.page_break()

    my_file.add('제목입니다')
    my_file.page_break()

    paragraphs = ['첫 문장 입니다. ' * 30,
                  '두번째 문장입니다. ' * 30,
                  '세번째 문장입니다. ' * 30]
    my_file.add(paragraphs)
    my_file.page_break()

    my_file.add([paragraphs[0]])
    my_file.blank_line()
    my_file.add([paragraphs[1]])
    my_file.page_break()

    my_file.add([paragraphs[0]])
    my_file.page_break()
    my_file.add([paragraphs[1]])
    my_file.page_break()

    
    my_file.add(image=icecream_png)
    my_file.page_break()

    df = pd.DataFrame([[10, 30, 50], [24, 13, 70]],
                      columns=['A사', 'B사', 'C사'],
                      index=['남', '여'])
    my_file.add(df)

    my_file.save('example.docx')


def test_doc_getting_started_pages2(bird_png, dog_png, horse_png, human_png):
    my_file = DOCX()  # 작업 시작

    # 1챕터
    my_file.add('동물들의 발자국')  # 제목

    rep = '의 발자국은 다음과 같이 생겼다.'

    my_file.add(['새{}'.format(rep)])  # 본문 내용 추가
    my_file.add(image=bird_png)  # 이미지 추가
    my_file.blank_line()  # 빈 줄 추가

    my_file.add(['개{}'.format(rep)])
    my_file.add(image=dog_png)
    my_file.blank_line()

    my_file.add(['말{}'.format(rep)])
    my_file.add(image=horse_png)
    my_file.blank_line()

    my_file.add(['사람{}'.format(rep)])
    my_file.add(image=human_png)

    my_file.page_break()  # 쪽 넘김

    # 2 챕터
    my_file.add('생김새 비교 표')
    df = pd.DataFrame([[2, 4, 4, 2],
                       ['O', 'X', 'X', 'X'],
                       ['O', 'O', 'O', 'X']],
                      columns=['새', '개', '말', '사람'],
                      index=['다리 수', '날개 유무', '꼬리 유무'])
    my_file.add(df)  # 표 추가
    my_file.blank_line()
    paragraphs = []  # 문단 만들기
    for col in df.columns:
        sentence = '- {name}는 {leg_num}개의 다리를 가지고 있으며 날개가 {has_wings}, 꼬리가 {has_tail}.'.format(
            name=col,
            leg_num=df[col].loc['다리 수'],
            has_wings='있고' if df[col].loc['날개 유무'] == 'O' else '없고',
            has_tail='있다' if df[col].loc['꼬리 유무'] == 'O' else '없다'
        )
        paragraphs.append(sentence)
    my_file.add(paragraphs)

    my_file.save('animals.docx')
