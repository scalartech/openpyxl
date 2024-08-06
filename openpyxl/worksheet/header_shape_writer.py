# Copyright (c) 2010-2023 openpyxl

import os
from typing import List
from openpyxl.xml.functions import (
    Element,
    SubElement,
    tostring,
)

from openpyxl.packaging.relationship import RelationshipList
from openpyxl.worksheet.header_footer import HeaderFooter, HeaderFooterItem
from openpyxl.worksheet.header_footer_adapter import HeaderFooterAdapter, ShapeMargins

vmlns = "urn:schemas-microsoft-com:vml"
officens = "urn:schemas-microsoft-com:office:office"
excelns = "urn:schemas-microsoft-com:office:excel"


class HeaderFooterShapeWriter:
    """
    Header Shape Writer for writing header shapes to the XML, uses legacyDrawingHF format.

    """

    shapes: List[ShapeMargins] = []
    hf: HeaderFooter

    def __init__(self, hf: HeaderFooter = None):
        self.hf = hf
        self.images = []
        self.rels = RelationshipList()
        self._add_header_footer_images()

    def _add_header_footer_images(self):
        self.images.clear()
        self.shapes.clear()

        for element in self.hf.__elements__:
            header_footer_item: HeaderFooterItem = getattr(self.hf, element)
            header_or_footer = "H" if "Header" in element else "F"
            adapter = HeaderFooterAdapter(self.images, header_footer_item, self.rels, header_or_footer)
            shapes = adapter.set_margins()
            self.shapes.extend(shapes)

    def write(self, root):

        if not hasattr(root, "findall"):
            root = Element("xml", nsmap={'v': vmlns, 'o': officens, 'x': excelns})

        for shape in self.shapes:
            title = os.path.splitext(os.path.basename(shape.image.ref))[0] if shape.image.ref is not None else "image"
            shape = _vml_image_shape_factory(shape.relationship.Id, title, shape.position, shape.image.height,
                                             shape.image.width)
            root.append(shape)

        return tostring(root)


def _vml_image_shape_factory(relationship, title, position, height, width):
    style = ("position:absolute;"
             "margin-left:0;"
             "margin-top:0;"
             "width:{width}pt;"
             "height:{height}pt;"
             "z-index:1;"
             "visibility:hidden").format(height=height,
                                         width=width)
    attrs = {
        "id": position,
        "{%s}spid" % officens: f"_x0000_s{ord(position[0])}{ord(position[1])}",
        "type": "#_x0000_t75",  # _x0000_t75 is an identifier for an image. # indicates an object, _x0000_ is a
        # compatibility prefix, and t75 specifies the shape type (image). This helps Excel track and place the image
        # in the worksheet correctly.
        "style": style
    }
    shape = Element("{%s}shape" % vmlns, attrs)

    SubElement(shape, "{%s}imagedata" % vmlns,
               {
                   "{%s}relid" % officens: str(relationship),
                   "{%s}title" % officens: title
               }
               )
    SubElement(shape, "{%s}lock" % officens,
               {"{%s}ext" % vmlns: "edit",
                "rotation": "t"})
    return shape
