from setuptools import setup

setup(
    name='firstaid',
    version='0.0.1',
    description='A First-Aid for Data Analysis and Reporting',
    author='QuantLab inc.',
    author_email='yu@mindscale.kr',
    packages=['firstaid'],
    install_requires=[
        'pandas',
        'Pillow',
        'python-pptx',
        'python-docx',
    ],
    package_data={
        '': ['*.txt', '*.md'],
        'firstaid': [],
    },
    entry_points={
    },
)
