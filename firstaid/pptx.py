from io import BytesIO

import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.util import Cm, Pt


class PPTX():
    def __init__(self):
        self.prs = Presentation()
        # dimensions
        self.SLIDE_HEIGHT = self.prs.slide_height.pt
        self.SLIDE_WIDTH = self.prs.slide_width.pt
        self.Y = Cm(4.45).pt
        self.X = Cm(1.27).pt
        self.TWO_ITEM_WIDTH = Cm(11.22).pt
        self.TWO_ITEM_HEIGHT = Cm(12.57).pt
        self.TWO_ITEM_X = Cm(12.91).pt
        self.ONE_ITEM_WIDTH = Cm(22.86).pt
        self.ONE_ITEM_HEIGHT = Cm(12.57).pt

    def add_bulletpoint(self, shapes, item, placeholder):
        body_shape = shapes.placeholders[placeholder]
        tf = body_shape.text_frame
        for i, row in enumerate(item):
            if i == 0:
                tf.text = row
            else:
                p = tf.add_paragraph()
                p.text = row
        return shapes

    def object_size(self, placeholder, layout_number, width=1000, height=1000):
        if placeholder == 1:
            if layout_number == 3:  # 2 item
                x = self.X
                width = min(width, self.TWO_ITEM_WIDTH)
                height = min(height, self.TWO_ITEM_HEIGHT)
            else:  # 1 item
                width = min(self.ONE_ITEM_WIDTH, width)
                height = min(self.ONE_ITEM_HEIGHT, height)
                x = max((self.SLIDE_WIDTH - width) / 2, self.X)
        else:
            x = self.TWO_ITEM_X
            width = min(width, self.TWO_ITEM_WIDTH)
            height = min(height, self.TWO_ITEM_HEIGHT)

        return x, width, height

    def add_image(self, shapes, item, placeholder, layout_number):
        im = Image.open(item)

        # image size
        width = im.width  # pt
        height = im.height  # pt

        x, width, height = self.object_size(placeholder, layout_number, width, height)

        # add image
        width, height = int(width), int(height)
        im = im.resize((width, height))

        fig = BytesIO()
        im.save(fig, 'PNG')
        fig.seek(0)
        shapes.add_picture(fig, left=Pt(x), top=Pt(self.Y), width=Pt(width), height=Pt(height))
        return shapes

    def add_table(self, shapes, df, placeholder, layout_number):
        x, width, height = self.object_size(placeholder, layout_number)
        table = shapes.add_table(df.shape[0] + 1, df.shape[1] + 1, Pt(x), Pt(self.Y), Pt(width), Pt(height)).table

        for i, col in enumerate(df.columns):
            table.cell(0, i + 1).text = str(col)

        for i, row in enumerate(df.index):
            table.cell(i + 1, 0).text = str(row)

        for i, row in enumerate(df.index):
            for e, item in enumerate(df.loc[row]):
                table.cell(i + 1, e + 1).text = str(item)

    def classify(self, shapes, content, placeholder, layout_number, placeholder_obj=None):
        if isinstance(content, list):
            return self.add_bulletpoint(shapes, content, placeholder)
        elif isinstance(content, str):
            return self.add_image(shapes, content, placeholder, layout_number)
        elif isinstance(content, pd.DataFrame):
            return self.add_table(shapes, content, placeholder, layout_number)

    def add(self, title, content1, content2=None):
        # slide layout
        try:
            layout_number = 3 if content2 else 1
        except ValueError:
            layout_number = 1 if content2.empty else 3
        # layout_number = 1 if content2 is None else 3
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[layout_number])

        # title
        shapes = slide.shapes
        shapes.title.text = title

        # content1
        placeholder = 1
        self.classify(shapes, content1, placeholder=placeholder, layout_number=layout_number)

        # content 2
        if layout_number == 3:
            placeholder = 2
            self.classify(shapes, content2, placeholder=placeholder, layout_number=layout_number)

    def save(self, file_name):
        self.prs.save(file_name)
        print('"{}" 파일이 생성되었습니다.'.format(file_name))
