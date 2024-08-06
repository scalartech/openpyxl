# Fixtures (pre-configured objects) for tests
import pytest
from openpyxl.drawing.image import Image
import os

@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    import os
    from py.path import local as LocalPath
    here = os.path.split(__file__)[0]
    DATADIR = os.path.join(here, "data")
    return LocalPath(DATADIR)

@pytest.fixture
def test_image():
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    EXPORT_IMAGE_PATH = os.path.join(BASE_DIR, f"tests/data/logo-excel.png".replace('/', os.sep))
    EXPORT_IMAGE = Image(EXPORT_IMAGE_PATH)
    return EXPORT_IMAGE