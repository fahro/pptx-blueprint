# TODO
from pptx_blueprint import Template, LibreOfficeNotFoundError
from pathlib import Path
import pytest
import os
import subprocess


@pytest.fixture
def template():
    filename = Path(__file__).absolute().parent / '../data/example01.pptx'
    tpl = Template(filename)
    return tpl


def test_open_template():
    filename = Path(__file__).absolute().parent / '../data/example01.pptx'
    tpl = Template(filename)


def test_open_template_missing():
    filename = Path(__file__).absolute().parent / '../data/non_existing.pptx'
    with pytest.raises(FileNotFoundError):
        tpl = Template(filename)


def test_find_shapes_from_all_slides(template):
    shapes = template._find_shapes('*', 'title')
    assert len(shapes) == 3
    for shape in shapes:
        assert shape.text == "{title}"


def test_find_shapes_from_one_slide(template):
    shapes = template._find_shapes(1, "logo")
    assert len(shapes) == 1
    assert shapes[0].text == '{logo}'


def test_find_shapes_index_out_of_range(template):
    with pytest.raises(IndexError):
        shapes = template._find_shapes(0, 'logo')


def test_save_pdf(template):
    try:
        subprocess.run(['libreoffice', '--version'],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)  # check if libreoffice is installed
        output_path = 'test/test.pdf'
        template.save_pdf(output_path)
        path = Path(output_path)
        assert path.exists() == True
        if path.exists():
            os.remove(path)
            Path.rmdir(path.parent)
    except FileNotFoundError:
        pytest.skip()
