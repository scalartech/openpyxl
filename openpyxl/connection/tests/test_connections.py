# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Connection():
    from ..connections import Connection
    return Connection


class TestConnection:


    def test_ctor(self, Connection):
        src_module = Connection(id=3, refreshedVersion=8, background=True, keepAlive=True)
        xml = tostring(src_module.to_tree())
        expected = """
        <connection id="3" keepAlive="1" refreshedVersion="8" background="1"
         interval="0" reconnectionMethod="1" minRefreshableVersion="0" savePassword="0"
         new="0" deleted="0" onlyUseConnectionFile="0" refreshOnLoad="0"
         saveData="0" credentials="integrated"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Connection):
        src = """
        <connection id="2" keepAlive="1"
        name="Query - Table1" description="Connection to the 'Table2' query in the workbook."
        type="5" refreshedVersion="8" background="1" saveData="1" />
        """
        node = fromstring(src)
        src_module = Connection.from_tree(node)
        assert src_module.saveData == True
        assert src_module.name == "Query - Table1"


@pytest.fixture
def DbPr():
    from ..connections import DbPr
    return DbPr


class TestDbPr:


    def test_ctor(self, DbPr):
        src_module = DbPr(connection="Data Model Connection", command="Model", commandType=True)
        xml = tostring(src_module.to_tree())
        expected = """
        <dbPr connection="Data Model Connection" command="Model" commandType="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DbPr):
        src = """
        <dbPr connection="Provider=Microsoft.Mashup.OleDb"
            command="SELECT * FROM [Table2]" />
        """
        node = fromstring(src)
        src_module = DbPr.from_tree(node)
        assert src_module.command == "SELECT * FROM [Table2]"


@pytest.fixture
def OlapPr():
    from ..connections import OlapPr
    return OlapPr


class TestOlapPr:


    def test_ctor(self, OlapPr):
        src_module = OlapPr(sendLocale=True, rowDrillCount=1000)
        xml = tostring(src_module.to_tree())
        expected = """
        <olapPr local="0" sendLocale="1" rowDrillCount="1000" localRefresh="1" serverFill="1" serverNumberFormat="1" serverFont="1" serverFontColor="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, OlapPr):
        src = """
        <olapPr serverFill="1" localRefresh="1" serverFontColor="1"/>
        """
        node = fromstring(src)
        src_module = OlapPr.from_tree(node)
        assert src_module == OlapPr(
            serverFill=True,
            localRefresh=True,
            serverFontColor=True
            )


@pytest.fixture
def TextField():
    from ..connections import TextField
    return TextField


class TestTextField:


    def test_ctor(self, TextField):
        src_module = TextField(type="text")
        xml = tostring(src_module.to_tree())
        expected = """
            <textField type="text" position="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TextField):
        from ..connections import TextField
        src = """
            <textField type="DMY" position="2" />
        """
        node = fromstring(src)
        src_module = TextField.from_tree(node)
        assert src_module == TextField(type="DMY", position=2)


@pytest.fixture
def TextPr():
    from ..connections import TextPr
    return TextPr


class TestTextPr:


    def test_ctor(self, TextPr):
        src_module = TextPr(prompt=False, codePage=437, sourceFile="C:\\Desktop\\text data.txt", delimiter="|")
        xml = tostring(src_module.to_tree())
        expected = """
        <textPr prompt="0" codePage="437" sourceFile="C:\\Desktop\\text data.txt" delimiter="|"
         fileType="win" firstRow="1" delimited="1" decimal="." thousands="," tab="1" qualifier="doubleQuote"
          space="0" comma="0" semicolon="0" consecutive="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TextPr):
        from ..connections import TextField
        src = """
        <textPr space="1" firstRow="13" sourceFile="C:\\Desktop\\text data.txt" delimiter="|">
            <textFields count="1">
                <textField />
            </textFields>
        </textPr>
        """
        node = fromstring(src)
        src_module = TextPr.from_tree(node)
        assert src_module == TextPr(space=True, firstRow=13, sourceFile="C:\\Desktop\\text data.txt", delimiter="|", textFields=[TextField()])


@pytest.fixture
def TableMissing():
    from ..connections import TableMissing
    return TableMissing


class TestTableMissing:


    @pytest.mark.xfail
    def test_ctor(self, TableMissing):
        src_module = TableMissing()
        xml = tostring(src_module.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.xfail
    def test_from_xml(self, TableMissing):
        src = """
        <root />
        """
        node = fromstring(src)
        src_module = TableMissing.from_tree(node)
        assert src_module == TableMissing()


@pytest.fixture
def Parameter():
    from ..connections import Parameter
    return Parameter


class TestParameter:


    def test_ctor(self, Parameter):
        src_module = Parameter(name="TestName", boolean=True, sqlType=4)
        xml = tostring(src_module.to_tree())
        expected = """
        <parameter name="TestName" boolean="1" sqlType="4" parameterType="prompt" refreshOnChange="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Parameter):
        src = """
        <parameter name="user" refreshOnChange="1" parameterType="cell" cell="Sheet1!$C$1"/>
        """
        node = fromstring(src)
        src_module = Parameter.from_tree(node)
        assert src_module == Parameter(name="user", refreshOnChange=True, parameterType="cell", cell="Sheet1!$C$1")


@pytest.fixture
def WebPr():
    from ..connections import WebPr
    return WebPr


class TestWebPr:


    def test_ctor(self, WebPr):
        src_module = WebPr(xml=True, firstRow=True, htmlTables=True)
        xml = tostring(src_module.to_tree())
        expected = """
        <webPr xml="1" firstRow="1" htmlTables="1" sourceData="0" parsePre="0" consecutive="0" xl97="0" textDates="0" xl2000="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WebPr):
        src = """
        <webPr sourceData="1" parsePre="1"
        consecutive="1" url="http://ServerName/" htmlTables="1" />
        """
        node = fromstring(src)
        src_module = WebPr.from_tree(node)
        assert src_module == WebPr(sourceData=True, parsePre=True, consecutive=True, url="http://ServerName/", htmlTables=True)


@pytest.fixture
def Tables():
    from ..connections import Tables
    return Tables


class TestTables:


    def test_ctor(self, Tables):
        src_module = Tables(x=3)
        xml = tostring(src_module.to_tree())
        expected = """
        <tables x="3" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Tables):
        src = """
        <tables s="test" />
        """
        node = fromstring(src)
        src_module = Tables.from_tree(node)
        assert src_module == Tables(s="test")


@pytest.fixture
def Connections():
    from ..connections import Connections
    return Connections


class TestConnections:


    def test_ctor(self, Connections):
        from ..connections import Connection
        src_module = Connections(connection=[Connection(id=1, refreshedVersion=4)])
        xml = tostring(src_module.to_tree())
        expected = """
        <connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <connection id="1" refreshedVersion="4" background="0" credentials="integrated" deleted="0"
            interval="0" keepAlive="0" minRefreshableVersion="0" new="0" onlyUseConnectionFile="0" reconnectionMethod="1" refreshOnLoad="0"
            saveData="0" savePassword="0"/>
        </connections>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Connections, datadir):
        datadir.chdir()
        with open("connections.xml", "rb") as src:
            node = fromstring(src.read())
        src_module = Connections.from_tree(node)
        assert len(src_module.connection) == 3
        assert src_module.connection[1].id == 2
        assert src_module.connection[2].saveData == True

