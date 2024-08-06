from openpyxl.worksheet.header_footer_adapter import HeaderFooterAdapter, ShapeMargins
from openpyxl.worksheet.header_shape_writer import HeaderFooterShapeWriter
from openpyxl.tests.helper import compare_xml
from openpyxl.workbook import Workbook
import pytest

@pytest.fixture
def _HeaderFooterPart():
    from ..header_footer import _HeaderFooterPart
    return _HeaderFooterPart

@pytest.fixture
def _HeaderFooterItem(_HeaderFooterPart, test_image):
    from ..header_footer import HeaderFooterItem
    part = _HeaderFooterPart(image=test_image, position="RH")
    item = HeaderFooterItem(left=part,right=part,center=None)

    return item


class TestHeaderFooterItem:

    def test_header_str(self, test_image):
        from ..header_footer import HeaderFooterItem
        header = HeaderFooterItem()
        header.left.image = test_image
        assert str(header) == "&L&G"

    def test_footer_str(self, test_image):
        from ..header_footer import HeaderFooterItem
        footer = HeaderFooterItem()
        footer.right.image = test_image
        assert str(footer) == "&R&G"


class TestHeaderFooterAdapter:

    def test_margins(self, _HeaderFooterItem):
        self.rels = []
        self.images = []
        self.header_footer_item = _HeaderFooterItem
        adapter = HeaderFooterAdapter(self.images, _HeaderFooterItem, self.rels, "H")
        margins = adapter._get_header_footer_info("left")

        assert isinstance(margins, ShapeMargins) is True
        assert margins.image.path == margins.relationship.target

    def test_shapes(self, _HeaderFooterItem):
        self.rels = []
        self.images = []
        adapter = HeaderFooterAdapter(self.images, _HeaderFooterItem, self.rels, "H")
        shapes = adapter.set_margins()
        print(shapes)
        assert len(shapes) == 2
        assert isinstance(shapes[0], ShapeMargins) is True
        assert isinstance(shapes[1], ShapeMargins) is True


class TestHeaderFooterShapeWriter:

    def test_shape_writer(self, _HeaderFooterItem, _HeaderFooterPart, test_image):
        from ..header_footer import HeaderFooterItem
        part = _HeaderFooterPart(image=test_image, position="RF")
        item = HeaderFooterItem(left=part, right=None, center=None, header_or_footer="F")

        wb = Workbook()
        ws = wb.create_sheet()
        ws.oddHeader = item

        w = HeaderFooterShapeWriter(ws.HeaderFooter)
        xml = w.write(None)

        expected = """
        <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:x="urn:schemas-microsoft-com:office:excel">
            <v:shape id="RF" o:spid="_x0000_s8270" type="#_x0000_t75"
                     style="position:absolute;margin-left:0;margin-top:0;width:1098pt;height:334pt;z-index:1;visibility:hidden">
                <v:imagedata o:relid="rId1" o:title="logo-excel"/>
                <o:lock v:ext="edit" rotation="t"/>
            </v:shape>
        </xml>"""

        diff = compare_xml(xml, expected)
        assert diff is None, diff
