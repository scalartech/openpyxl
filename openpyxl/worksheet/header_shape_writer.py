# Copyright (c) 2010-2023 openpyxl

from xml.etree.ElementTree import QName
from openpyxl.xml.functions import (
    Element,
    SubElement,
    tostring,
)

from openpyxl.drawing.image import Image

vmlns = "urn:schemas-microsoft-com:vml"
officens = "urn:schemas-microsoft-com:office:office"
excelns = "urn:schemas-microsoft-com:office:excel"


class HeaderShapeWriter(object):

    images = []

    def add_header_image(self, image):
        self.images.append(image)

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


    def add_vml_image_shape(self, root, idx, position, height, width):
        shape = _vml_image_shape_factory(idx, "image", position, height, width)

        # shape.set('id', "_x0000_s%04d" % idx)
        root.append(shape)


    def write(self, root):

        if not hasattr(root, "findall"):
            root = Element("xml", nsmap={'v': vmlns, 'o': officens, 'x': excelns})

        # check whether image shape type already exists
        shape_types = root.find("{%s}shapetype[@id='_x0000_t75']" % vmlns)
        if shape_types is None:
            self.add_vml_image_shapetype(root)

        for image in self.images:
            rel_id = "rId%s" % image._id
            self.add_vml_image_shape(root, rel_id, image.anchor, image.height, image.width)

        return tostring(root)


def _vml_image_shape_factory(relationship, title, position, height, width):
    style = ("position:absolute; "
             "margin-left:0;"
             "margin-top:0;"
             "width:{width}pt;"
             "height:{height}pt;"
             "z-index:1;"
             "visibility:hidden").format(height=height,
                                         width=width)
    attrs = {
        "id": position,
        "{%s}spid" % officens : "_x0000_s1025",
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
