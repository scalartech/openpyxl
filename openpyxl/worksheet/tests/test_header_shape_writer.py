# Copyright (c) 2010-2023 openpyxl


import pytest

from openpyxl.drawing.image import Image
from openpyxl.workbook import Workbook
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import (
    fromstring,
    tostring,
    Element,
)

from openpyxl.worksheet.header_shape_writer import (
    HeaderShapeWriter,
    vmlns,
    excelns,
    officens
)

class TestHeaderImages:

    @pytest.mark.pil_required
    def test_add_header_image(self, datadir):
        datadir.chdir()

        # User image from Image tests
        header_image = Image('../../../writer/tests/data/plain.png')
        hw = HeaderShapeWriter(header_image)

        content = fromstring(hw.write(None))

        shape = content.find('{%s}shape' % vmlns)

        assert len(content.findall('{%s}shapetype' % vmlns)) == 1
        assert len(content.findall('{%s}shapelayout' % officens)) == 1
        assert shape.attrib['style'] == "position:absolute; margin-left:0;margin-top:0;width:118pt;height:118pt;z-index:1;visibility:hidden"

