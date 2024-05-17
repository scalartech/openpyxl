# Copyright (c) 2010-2023 openpyxl

import os
from typing import List
from openpyxl.drawing.image import Image
from openpyxl.xml.functions import (
    Element,
    SubElement,
    tostring,
)
from openpyxl.descriptors import (
    Alias,
    Bool,
    Strict,
    String,
    Integer,
    MatchPattern,
    Typed,
)
from openpyxl.packaging.relationship import (
    Relationship,
    RelationshipList,
)
from openpyxl.worksheet.header_footer import HeaderFooter, HeaderFooterItem, _HeaderFooterPart

vmlns = "urn:schemas-microsoft-com:vml"
officens = "urn:schemas-microsoft-com:office:office"
excelns = "urn:schemas-microsoft-com:office:excel"

HEADER_IMAGE_ANCHORS = ["LH", "CH", "RH", "LF", "CF", "RF"]

class HeaderShape(Strict):
    """
    Header Shape that includes an `image`, the `anchor` or position in the header or footer, and the `relationship` for this shape.

    """
    
    image = Typed(expected_type=Image, allow_none=False)
    anchor = String(allow_none=False)
    relationship = Typed(expected_type=Relationship, allow_none=False)
    
    def __init__(self, image, anchor, relationship):
        self.image = image
        self.anchor = anchor
        self.relationship = relationship


class HeaderShapeWriter(object):
    """
    Header Shape Writer for writing header shapes to the XML, uses legacyDrawingHF format.

    """

    shapes: List[HeaderShape] = []
    hf: HeaderFooter
    images = []
    rels = RelationshipList()

    def __init__(self, hf:HeaderFooter):
        self.hf = hf
        self._add_header_images()

    def _add_header_images(self):
        for element in self.hf.__elements__:
            header_footer_item: HeaderFooterItem = getattr(self.hf, element)
            header_or_footer = "H" if "Header" in element else "F"
            self._process_header_footer_item(header_footer_item, header_or_footer)
    
    def _process_header_footer_item(self, header_footer_item, header_or_footer):
        for key in ("left", "center", "right"):
            header_footer_part: _HeaderFooterPart = getattr(header_footer_item, key)
            header_image = header_footer_part.image
            if header_image is not None:
                # Add image to list of images to save in Workbook
                if header_image not in self.images:
                    self.images.append(header_image)
                    header_image._id = len(self.images)
                
                # Check if Relationship with Target already exists in rels
                rel = next(
                    (check_rel for check_rel in self.rels.Relationship if check_rel.Target == header_image.path), 
                    Relationship(type="image", Target=header_image.path)
                )
                if rel not in self.rels.Relationship:
                    self.rels.append(rel)

                # Determine anchor position, such as "CH" for center header
                anchor = key[0].upper() + header_or_footer
                if anchor not in HEADER_IMAGE_ANCHORS:
                    raise ValueError("Invalid header image anchor position, must be one of %s" % HEADER_IMAGE_ANCHORS)
                header_shape = HeaderShape(header_image, anchor, rel)
                self.shapes.append(header_shape)


    def add_vml_image_shapetype(self, root):
        shape_layout = SubElement(root, "{%s}shapelayout" % officens,
                                  {"{%s}ext" % vmlns: "edit"})
        SubElement(shape_layout,
                   "{%s}idmap" % officens,
                   {"{%s}ext" % vmlns: "edit", "data": "1"})
        shape_type = SubElement(root,
                                "{%s}shapetype" % vmlns,
                                {"id": "_x0000_t75",
                                 "coordsize": "21600,21600",
                                 "{%s}spt" % officens: "75",
                                 "{%s}preferrelative" % officens: "t",
                                 "path": "m@4@5l@4@11@9@11@9@5xe",
                                 "filled": "f",
                                 "stroked": "f"})
        SubElement(shape_type, "{%s}stroke" % vmlns, {"joinstyle": "miter"})
        formulas = SubElement(shape_type, "{%s}formulas" % vmlns)
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "if lineDrawn pixelLineWidth 0"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "sum @0 1 0"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "sum 0 0 @1"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "prod @2 1 2"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "prod @3 21600 pixelWidth"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "prod @3 21600 pixelHeight"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "sum @0 0 1"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "prod @6 1 2"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "prod @7 21600 pixelWidth"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "sum @8 21600 0"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "prod @7 21600 pixelHeight"})
        SubElement(formulas, "{%s}f" % vmlns, {"eqn": "sum @10 21600 0"})
        SubElement(shape_type,
                   "{%s}path" % vmlns,
                   {"{%s}extrusionok" % officens: "f",
                    "gradientshapeok": "t",
                    "{%s}connecttype" % officens: "rect"})
        SubElement(shape_type,
                   "{%s}lock" % officens,
                   {"{%s}ext" % vmlns: "edit",
                    "aspectratio": "t"})


    def write(self, root):

        if not hasattr(root, "findall"):
            root = Element("xml", nsmap={'v': vmlns, 'o': officens, 'x': excelns})

        # check whether image shape type already exists
        shape_types = root.find("{%s}shapetype[@id='_x0000_t75']" % vmlns)
        if shape_types is None:
            self.add_vml_image_shapetype(root)

        for shape in self.shapes:
            title = os.path.splitext(os.path.basename(shape.image.ref))[0] if shape.image.ref is not None else "image"
            shape = _vml_image_shape_factory(shape.relationship.Id, title, shape.anchor, shape.image.height, shape.image.width)
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
        "{%s}spid" % officens : f"_x0000_s{ord(position[0])}{ord(position[1])}",
        "type": "#_x0000_t75",
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
