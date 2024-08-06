from openpyxl.packaging.relationship import Relationship
from openpyxl.worksheet.header_footer import _HeaderFooterPart
from openpyxl.descriptors import (
    Strict,
    String,
    Typed,
)
from openpyxl.drawing.image import Image


class ShapeMargins(Strict):
    """
    Header/Footer Shape that includes an `image`, the `anchor` or position in the header or footer, and the `relationship` for this shape.

    """

    image = Typed(expected_type=Image, allow_none=False)
    position = String(allow_none=False)
    relationship = Typed(expected_type=Relationship, allow_none=False)

    def __init__(self, image, position, relationship):
        self.image = image
        self.position = position
        self.relationship = relationship


class HeaderFooterAdapter:
    """
    Adapter for HeaderFooterItem to provide necessary details for HeaderFooterShapeWriter.
    """

    def __init__(self, images, header_footer_item, relationships, header_or_footer):
        self.header_footer_item = header_footer_item
        self.rels = relationships
        self.header_or_footer = header_or_footer
        self.images = images

    def set_margins(self):
        shapes = []
        for position in ("left", "center", "right"):
            shape_info = self._get_header_footer_info(position)
            if shape_info:
                shapes.append(shape_info)
        return shapes

    def _get_header_footer_info(self, position):
        header_footer_part: _HeaderFooterPart = getattr(self.header_footer_item, position)
        header_image = header_footer_part.image
        if header_image is not None:
            if header_image not in self.images:
                self.images.append(header_image)
                header_image._id = len(self.images)

            rel = self._get_or_create_relationship(header_image)
            position = header_footer_part.position
            return ShapeMargins(header_image, position, rel)

    def _get_or_create_relationship(self, header_image):
        for check_rel in self.rels:
            if check_rel.Target == header_image.path:
                return check_rel

        # If not found, create a new relationship
        rel = Relationship(type="image", Target=header_image.path)
        self.rels.append(rel)
        return rel
