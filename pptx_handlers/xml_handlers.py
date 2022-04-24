from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls


def sub_element(parent, tagname, **kwargs):  # necessary for table formatting
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element


def page_number_xml():
    fld_xml = (
            '<a:fld %s id="{1F4E2DE4-8ADA-4D4E-9951-90A1D26586E7}" type="slidenum">\n'
            '  <a:rPr lang="en-US" smtClean="0"/>\n'
            '  <a:t>2</a:t>\n'
            '</a:fld>\n' % nsdecls("a")
    )
    fld = parse_xml(fld_xml)
    return fld


def add_table_line_graph(chart):
    plotArea = chart._element.chart.plotArea
    sub_element(plotArea, 'c:dTable')  # Add data table tag
    dataTable = chart._element.chart.plotArea.dTable
    # Add values to data table tag
    for sub in ['c:showHorzBorder', 'c:showVertBorder', 'c:showOutline', 'c:showKeys']:
        sub_element(dataTable, sub, val='1')


def remove_cell_borders_with_xml(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for lines in ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']:
        ln = sub_element(tcPr, lines, w='0', cap='flat', cmpd='sng', algn='ctr')

