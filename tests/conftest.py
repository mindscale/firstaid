from pathlib import Path

import pytest


curdir = Path(__file__).parent


def png(filename):
    return str((curdir / '{}.PNG'.format(filename)).absolute())


@pytest.fixture
def bird_png():
    return png('bird')


@pytest.fixture
def coffee_png():
    return png('coffee')


@pytest.fixture
def dog_png():
    return png('dog')


@pytest.fixture
def icecream_png():
    return png('icecream')


@pytest.fixture
def horse_png():
    return png('horse')


@pytest.fixture
def human_png():
    return png('human')
