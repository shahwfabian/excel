"""Add the two Lab 3 PivotCharts (cMonthlyRevenue, cRevenueBySalesManager) to
lab03.xlsx by editing the OOXML package directly, leaving the Power Pivot
Data Model, queries and the Lab 2 charts untouched."""
import json, re, sys, zipfile
from datetime import datetime
from xml.sax.saxutils import escape

SRC, AGG, DST = sys.argv[1], sys.argv[2], sys.argv[3]
BOOK = DST.replace('\\', '/').split('/')[-1]
agg = json.load(open(AGG, encoding='utf-8'))
zin = zipfile.ZipFile(SRC)
parts = {n: zin.read(n) for n in zin.namelist()}
txt = lambda n: parts[n].decode('utf-8')
def put(n, s): parts[n] = s.encode('utf-8')

REGIONS = agg['regions']
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
XR = 'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"'

# ---------------------------------------------------------------- pivot caches
cache1 = txt('xl/pivotCache/pivotCacheDefinition1.xml')
hier_block = re.search(r'<cacheHierarchies.*</cacheHierarchies>', cache1).group(0)
tail = re.search(r'<kpis.*?</maps>', cache1).group(0)
head = cache1[:cache1.index('<cacheSource')]
hiers = re.findall(r'<cacheHierarchy [^>]*?/>|<cacheHierarchy [^>]*[^/]>.*?</cacheHierarchy>', hier_block)
assert len(hiers) == 29

def plain_hier(h):
    """Reset a cache hierarchy to 'unused' (as Excel writes it)."""
    h = re.sub(r'<fieldsUsage.*?</fieldsUsage>', '', h)
    h = re.sub(r' count="\d+"', ' count="0"', h)
    if h.endswith('></cacheHierarchy>'):
        h = h[:-len('></cacheHierarchy>')] + '/>'
    return h

def used_hier(h, field_idx):
    h = plain_hier(h)[:-2]
    h = h.replace(' count="0"', ' count="2"')
    return h + f'><fieldsUsage count="2"><fieldUsage x="-1"/><fieldUsage x="{field_idx}"/></fieldsUsage></cacheHierarchy>'

measure = ('<cacheHierarchy uniqueName="[Measures].[Sum of Revenue]" caption="Sum of Revenue" measure="1" '
           'displayFolder="" measureGroup="Sales" count="0" oneField="1" hidden="1"><fieldsUsage count="1">'
           '<fieldUsage x="{x}"/></fieldsUsage><extLst><ext uri="{{B97F6D7D-B522-45F9-BDA1-12C45D357490}}" '
           'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:cacheHierarchy '
           'aggregatedColumn="13"/></ext></extLst></cacheHierarchy>')

def cache_xml(uid, cache_id, fields, used):
    hs = []
    for i, h in enumerate(hiers):
        if i == 28:
            hs.append(measure.format(x=len(fields)))
        elif i in used:
            hs.append(used_hier(h, used[i]))
        else:
            hs.append(plain_hier(h))
    fields = fields + ['<cacheField name="[Measures].[Sum of Revenue]" caption="Sum of Revenue" numFmtId="0" hierarchy="28" level="32767"/>']
    h = re.sub(r'refreshedDate="[^"]*"', 'refreshedDate="46292.5"', head)
    h = re.sub(r'xr:uid="[^"]*"', f'xr:uid="{uid}"', h)
    h = h.replace('saveData="0"', 'saveData="0" refreshOnLoad="1"')
    return (h + '<cacheSource type="external" connectionId="5"><extLst><ext uri="{F057638F-6D5F-4e77-A914-E7F072B9BCA8}" '
            'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sourceConnection name="ThisWorkbookDataModel"/>'
            f'</ext></extLst></cacheSource><cacheFields count="{len(fields)}">' + ''.join(fields) + '</cacheFields>'
            f'<cacheHierarchies count="29">{"".join(hs)}</cacheHierarchies>' + tail +
            '<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
            f'<x14:pivotCacheDefinition pivotCacheId="{cache_id}" supportSubqueryNonVisual="1" supportSubqueryCalcMem="1" supportAddCalcMems="1"/></ext>'
            '<ext uri="{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
            '<x15:pivotCacheIdVersion cacheIdSupportedVersion="6" cacheIdCreatedVersion="7"/></ext></extLst></pivotCacheDefinition>')

def str_field(name, caption, hier, values):
    return (f'<cacheField name="{name}" caption="{caption}" numFmtId="0" hierarchy="{hier}" level="1">'
            f'<sharedItems count="{len(values)}">' + ''.join(f'<s v="{escape(v)}"/>' for v in values) + '</sharedItems></cacheField>')

months = agg['months']
iso = [m + 'T00:00:00' for m in months]
month_field = (
    '<cacheField name="[Dates].[Start of Month].[Start of Month]" caption="Start of Month" numFmtId="0" hierarchy="2" level="1">'
    f'<sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="{iso[0]}" '
    f'maxDate="{iso[-1]}" count="{len(iso)}">' + ''.join(f'<d v="{d}"/>' for d in iso) + '</sharedItems>'
    '<extLst><ext uri="{4F2E5C28-24EA-4eb8-9CBF-B6C8F9C3D259}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
    '<x15:cachedUniqueNames>' + ''.join(
        f'<x15:cachedUniqueName index="{i}" name="[Dates].[Start of Month].&amp;[{d}]"/>' for i, d in enumerate(iso)) +
    '</x15:cachedUniqueNames></ext></extLst></cacheField>')
region_field = str_field('[Region].[Region].[Region]', 'Region', 5, REGIONS)
MANAGERS = agg['managers']  # alphabetical, as the OLAP member order
mgr_field = str_field('[SalesManager].[Sales Manager].[Sales Manager]', 'Sales Manager', 21, MANAGERS)

CID3, CID4 = 1734905611, 1402377920
put('xl/pivotCache/pivotCacheDefinition3.xml',
    cache_xml('{7C3B1F0A-5E2D-4B8A-9C61-3F0D2A7E4B11}', CID3, [month_field, region_field], {2: 0, 5: 1}))
put('xl/pivotCache/pivotCacheDefinition4.xml',
    cache_xml('{2E9A6D4C-81B7-4F35-A0D2-6C5B9E1F3A22}', CID4, [mgr_field, region_field], {21: 0, 5: 1}))

# ---------------------------------------------------------------- pivot tables
pt1 = txt('xl/pivotTables/pivotTable1.xml')
phier = re.search(r'<pivotHierarchies.*</pivotHierarchies>', pt1).group(0)

def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def num(v):
    return repr(float(v))

def pivot_xml(name, uid, cache_idx, cache_id, row_items_order, row_hier, row_vals, col_vals, data, sort_by_value, active):
    """row_items_order: cache indexes of the row field in display order.
    data[col][cacheRowIdx] -> value or None."""
    nr, nc = len(row_items_order), len(col_vals)
    ref = f'A1:{col_letter(nc + 2)}{nr + 3}'
    sort_attr = ' sortType="ascending"' if sort_by_value else ''
    auto = ('<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1">'
            '<reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea></autoSortScope>') if sort_by_value else ''
    f0 = (f'<pivotField axis="axisRow" allDrilled="1" subtotalTop="0" showAll="0"{sort_attr} dataSourceSort="1" defaultSubtotal="0" '
          f'defaultAttributeDrillState="1"><items count="{nr}">' + ''.join(f'<item x="{i}"/>' for i in row_items_order) + f'</items>{auto}</pivotField>')
    f1 = (f'<pivotField axis="axisCol" allDrilled="1" subtotalTop="0" showAll="0" dataSourceSort="1" defaultSubtotal="0" '
          f'defaultAttributeDrillState="1"><items count="{nc}">' + ''.join(f'<item x="{i}"/>' for i in range(nc)) + '</items></pivotField>')
    f2 = '<pivotField dataField="1" subtotalTop="0" showAll="0" defaultSubtotal="0"/>'
    def items(n):
        return ''.join('<i><x/></i>' if i == 0 else f'<i><x v="{i}"/></i>' for i in range(n)) + '<i t="grand"><x/></i>'
    fmts = ''.join(
        f'<chartFormat chart="0" format="{k}" series="1"><pivotArea type="data" outline="0" fieldPosition="0"><references count="2">'
        f'<reference field="4294967294" count="1" selected="0"><x v="0"/></reference>'
        f'<reference field="1" count="1" selected="0"><x v="{k}"/></reference></references></pivotArea></chartFormat>'
        for k in range(nc))
    # cached cell values (data area incl. grand totals) for x15:pivotTableData
    rows = []
    for ci in row_items_order:
        vals = [data[c][ci] for c in col_vals]
        cells = ''.join('<x15:c t="bl"/>' if v is None else f'<x15:c><x15:v>{num(v)}</x15:v></x15:c>' for v in vals)
        cells += f'<x15:c><x15:v>{num(sum(v for v in vals if v is not None))}</x15:v></x15:c>'
        rows.append(f'<x15:pivotRow count="{nc + 1}">{cells}</x15:pivotRow>')
    tots = [sum(v for v in data[c] if v is not None) for c in col_vals]
    rows.append(f'<x15:pivotRow count="{nc + 1}">' + ''.join(f'<x15:c><x15:v>{num(v)}</x15:v></x15:c>' for v in tots + [sum(tots)]) + '</x15:pivotRow>')
    ui = ''.join(f'<x15:activeTabTopLevelEntity name="[{e}]"/>' for e in active)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" {XR} xr:uid="{uid}" '
            f'name="{name}" cacheId="{cache_idx}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
            'applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" '
            'useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0" chartFormat="1">'
            f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>'
            f'<pivotFields count="3">{f0}{f1}{f2}</pivotFields>'
            f'<rowFields count="1"><field x="0"/></rowFields><rowItems count="{nr + 1}">{items(nr)}</rowItems>'
            f'<colFields count="1"><field x="1"/></colFields><colItems count="{nc + 1}">{items(nc)}</colItems>'
            '<dataFields count="1"><dataField name="Sum of Revenue" fld="2" baseField="0" baseItem="0"/></dataFields>'
            f'<chartFormats count="{nc}">{fmts}</chartFormats>{phier}'
            f'<rowHierarchiesUsage count="1"><rowHierarchyUsage hierarchyUsage="{row_hier}"/></rowHierarchiesUsage>'
            '<colHierarchiesUsage count="1"><colHierarchyUsage hierarchyUsage="5"/></colHierarchiesUsage>'
            '<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
            '<x14:pivotTableDefinition calculatedMembersInFilters="1" hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext>'
            '<ext uri="{44433962-1CF7-4059-B4EE-95C3D5FFCF73}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
            f'<x15:pivotTableData rowCount="{nr + 1}" columnCount="{nc + 1}" cacheId="{cache_id}">{"".join(rows)}</x15:pivotTableData></ext>'
            '<ext uri="{E67621CE-5B39-4880-91FE-76760E9C1902}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
            f'<x15:pivotTableUISettings>{ui}</x15:pivotTableUISettings></ext>'
            '<ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout">'
            '<xpdl:pivotTableDefinition16 EnabledSubtotalsDefault="0" SubtotalsOnTopDefault="0"/></ext></extLst></pivotTableDefinition>')

# Revenue by Month: rows = Start of Month (chronological), columns = Region
put('xl/pivotTables/pivotTable3.xml', pivot_xml(
    'PivotChartTable3', '{5B8E2C71-0F3A-4D69-8E24-9A1C7D3B6F33}', 2, CID3,
    list(range(len(months))), 2, months, REGIONS, agg['monthly'], False, ['Dates', 'Region', 'Sales']))

# Revenue by Sales Manager: rows = Sales Manager sorted ascending by Sum of Revenue
mgr_tot = [sum(agg['bymgr'][r][i] or 0 for r in REGIONS) for i in range(len(MANAGERS))]
mgr_order = sorted(range(len(MANAGERS)), key=lambda i: mgr_tot[i])
put('xl/pivotTables/pivotTable4.xml', pivot_xml(
    'PivotChartTable4', '{A4D7F390-6C1E-4B2F-B8A5-0E7C94D2F144}', 3, CID4,
    mgr_order, 21, MANAGERS, REGIONS, agg['bymgr'], True, ['SalesManager', 'Region', 'Sales']))
rel = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
       'Target="../pivotCache/pivotCacheDefinition{}.xml"/></Relationships>')
put('xl/pivotTables/_rels/pivotTable3.xml.rels', rel.format(3))
put('xl/pivotTables/_rels/pivotTable4.xml.rels', rel.format(4))

# ---------------------------------------------------------------- charts
GRAY = '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>'
FONT = '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>'
NOLINE = '<a:ln><a:noFill/></a:ln>'
GRID = ('<c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1">'
        '<a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:majorGridlines>')
AXLINE = ('<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/>'
          '<a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>')

def txpr(sz, rot='-60000000'):
    return (f'<c:txPr><a:bodyPr rot="{rot}" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>'
            f'<a:lstStyle/><a:p><a:pPr><a:defRPr sz="{sz}" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">{GRAY}{FONT}'
            '</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>')

def title(text):
    return ('<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>'
            f'<a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">{GRAY}{FONT}'
            f'</a:defRPr></a:pPr><a:r><a:rPr lang="en-US"/><a:t>{escape(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/>'
            f'<c:spPr><a:noFill/>{NOLINE}<a:effectLst/></c:spPr>'
            '<c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>'
            f'<a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">{GRAY}{FONT}'
            '</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:title><c:autoTitleDeleted val="0"/>')

def ser_sppr(k, line):
    fill = f'<a:solidFill><a:schemeClr val="accent{k + 1}"/></a:solidFill>'
    if line:
        return f'<c:spPr><a:ln w="28575" cap="rnd">{fill}<a:round/></a:ln><a:effectLst/></c:spPr>'
    return f'<c:spPr>{fill}{NOLINE}<a:effectLst/></c:spPr>'

def pivot_fmts(n, line):
    marker = '<c:marker><c:symbol val="none"/></c:marker>'
    return '<c:pivotFmts>' + ''.join(
        f'<c:pivotFmt><c:idx val="{k}"/>{ser_sppr(k, line)}{marker}</c:pivotFmt>' for k in range(n)) + '</c:pivotFmts>'

def series(k, name, cats, vals, line, tag):
    cat = f'<c:cat><c:strLit><c:ptCount val="{len(cats)}"/>' + ''.join(
        f'<c:pt idx="{i}"><c:v>{escape(c)}</c:v></c:pt>' for i, c in enumerate(cats)) + '</c:strLit></c:cat>'
    val = f'<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="{len(vals)}"/>' + ''.join(
        f'<c:pt idx="{i}"><c:v>{num(v)}</c:v></c:pt>' for i, v in enumerate(vals) if v is not None) + '</c:numLit></c:val>'
    extra = '<c:marker><c:symbol val="none"/></c:marker>' if line else '<c:invertIfNegative val="0"/>'
    end = '<c:smooth val="0"/>' if line else ''
    return (f'<c:ser><c:idx val="{k}"/><c:order val="{k}"/><c:tx><c:v>{escape(name)}</c:v></c:tx>{ser_sppr(k, line)}{extra}{cat}{val}{end}'
            '<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
            f'<c16:uniqueId val="{{0000000{k}-{tag}}}"/></c:ext></c:extLst></c:ser>')

def cat_ax(ax, cross, pos, lbl_rot, skip=''):
    return (f'<c:catAx><c:axId val="{ax}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{pos}"/>'
            '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            f'<c:spPr><a:noFill/>{AXLINE}<a:effectLst/></c:spPr>{txpr(900, lbl_rot)}<c:crossAx val="{cross}"/><c:crosses val="autoZero"/>'
            f'<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>{skip}<c:noMultiLvlLbl val="0"/></c:catAx>')

def val_ax(ax, cross, pos):
    return (f'<c:valAx><c:axId val="{ax}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{pos}"/>{GRID}'
            '<c:numFmt formatCode="#,,&quot;M&quot;" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
            f'<c:tickLblPos val="nextTo"/><c:spPr><a:noFill/>{NOLINE}<a:effectLst/></c:spPr>{txpr(900)}<c:crossAx val="{cross}"/>'
            '<c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')

def chart_xml(ttl, plot, n, line, pivot_name):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<c:chartSpace xmlns:c="{C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">'
            '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>'
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" '
            'xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/>'
            f'</mc:Fallback></mc:AlternateContent><c:chart>{title(ttl)}{pivot_fmts(n, line)}<c:plotArea><c:layout/>{plot}'
            f'<c:spPr><a:noFill/>{NOLINE}<a:effectLst/></c:spPr></c:plotArea>'
            # Legend: Top, 12 pt
            f'<c:legend><c:legendPos val="t"/><c:overlay val="0"/><c:spPr><a:noFill/>{NOLINE}<a:effectLst/></c:spPr>{txpr(1200, "0")}</c:legend>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" '
            'xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/>'
            '</c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart>'
            # Chart area: white fill, Shape Outline = No Outline
            f'<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:noFill/><a:round/></a:ln>'
            '<a:effectLst/></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
            '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>'
            '<c:extLst><c:ext uri="{723BEF56-08C2-4564-9609-F4CBC75E7E54}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">'
            f'<c15:pivotSource><c15:name>[{BOOK}]{pivot_name}</c15:name><c15:fmtId val="0"/></c15:pivotSource>'
            # Field Buttons > Hide All (no dropZonesVisible element)
            '<c15:pivotOptions><c15:dropZoneFilter val="1"/><c15:dropZoneCategories val="1"/><c15:dropZoneData val="1"/><c15:dropZoneSeries val="1"/>'
            '</c15:pivotOptions></c:ext><c:ext uri="{E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
            '<c16:pivotOptions16><c16:showExpandCollapseFieldButtons val="1"/></c16:pivotOptions16></c:ext></c:extLst></c:chartSpace>')

# Line chart - Revenue by Month
mcats = [f'{int(m[5:7])}/{int(m[8:10])}/{m[:4]}' for m in months]
ser_line = ''.join(series(k, r, mcats, agg['monthly'][r], True, '3C1D-4E2A-9B7F-5A6C8D2E1F03') for k, r in enumerate(REGIONS))
plot_line = (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{ser_line}'
             '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
             '<c:showBubbleSize val="0"/></c:dLbls><c:smooth val="0"/><c:axId val="1502843601"/><c:axId val="1502845041"/></c:lineChart>'
             + cat_ax(1502843601, 1502845041, 'b', '-5400000') + val_ax(1502845041, 1502843601, 'l'))
put('xl/charts/chart3.xml', chart_xml('Revenue by Month', plot_line, len(REGIONS), True, 'PivotChartTable3'))

# Stacked bar - Revenue by Sales Manager (categories in the pivot's ascending order)
bcats = [MANAGERS[i] for i in mgr_order]
ser_bar = ''.join(series(k, r, bcats, [agg['bymgr'][r][i] for i in mgr_order], False, '7B2E-4C9D-A1F6-3E8B5D0C2A04')
                  for k, r in enumerate(REGIONS))
plot_bar = (f'<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>{ser_bar}'
            '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
            '<c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="150"/><c:overlap val="100"/><c:axId val="1611237841"/><c:axId val="1611239281"/></c:barChart>'
            + cat_ax(1611237841, 1611239281, 'l', '-60000000', '<c:tickLblSkip val="1"/>') + val_ax(1611239281, 1611237841, 'b'))
put('xl/charts/chart4.xml', chart_xml('Revenue by Sales Manager', plot_bar, len(REGIONS), False, 'PivotChartTable4'))

for n in (3, 4):
    parts[f'xl/charts/style{n}.xml'] = parts['xl/charts/style1.xml']
    parts[f'xl/charts/colors{n}.xml'] = parts['xl/charts/colors1.xml']
    put(f'xl/charts/_rels/chart{n}.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors{n}.xml"/>'
        f'<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style{n}.xml"/></Relationships>')

# Existing Lab 2 charts point at the old file name; keep them consistent with the saved name.
for n in (1, 2):
    put(f'xl/charts/chart{n}.xml', txt(f'xl/charts/chart{n}.xml').replace('[lab03.xlsx]', f'[{BOOK}]'))

# ---------------------------------------------------------------- drawing (dashboard placement)
def frame(fr, to, cid, name, rid, guid):
    a = lambda c, co, r, ro: f'<xdr:col>{c}</xdr:col><xdr:colOff>{co}</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>{ro}</xdr:rowOff>'
    return (f'<xdr:twoCellAnchor><xdr:from>{a(*fr)}</xdr:from><xdr:to>{a(*to)}</xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
            f'<xdr:cNvPr id="{cid}" name="{name}"><a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}">'
            f'<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{guid}"/></a:ext></a:extLst></xdr:cNvPr>'
            '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>'
            f'<a:graphicData uri="{C}"><c:chart xmlns:c="{C}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/>'
            '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')

d = txt('xl/drawings/drawing1.xml')
# cMonthlyRevenue: right of "Revenue by Year", under the Revenue / Total Cost cards (F4:H18)
# cRevenueBySalesManager: right-hand column beside "Revenue by Order Method", tall enough for all 14 managers (K19:Q36)
d = d.replace('</xdr:wsDr>',
              frame((5, 4762, 3, 133351), (8, 2352675, 17, 123826), 5, 'cMonthlyRevenue', 'rId4', '{9E3F1B27-4A6C-D0E5-8B21-7C4A3E9F0D15}') +
              frame((10, 4762, 18, 57150), (16, 0, 36, 0), 6, 'cRevenueBySalesManager', 'rId5', '{3B7D5E91-C2A8-4F60-1D3E-A95B7C2E8F46}') +
              '</xdr:wsDr>')
put('xl/drawings/drawing1.xml', d)
put('xl/drawings/_rels/drawing1.xml.rels', txt('xl/drawings/_rels/drawing1.xml.rels').replace('</Relationships>',
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>'
    '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart4.xml"/></Relationships>'))

# ---------------------------------------------------------------- workbook wiring
wr = txt('xl/_rels/workbook.xml.rels')
wr = wr.replace('</Relationships>',
    '<Relationship Id="rId31" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition3.xml"/>'
    '<Relationship Id="rId32" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition4.xml"/>'
    '<Relationship Id="rId33" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="pivotTables/pivotTable3.xml"/>'
    '<Relationship Id="rId34" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="pivotTables/pivotTable4.xml"/></Relationships>')
put('xl/_rels/workbook.xml.rels', wr)
wb = txt('xl/workbook.xml')
wb = wb.replace('<pivotCache cacheId="1" r:id="rId4"/></x15:pivotCaches>',
                '<pivotCache cacheId="1" r:id="rId4"/><pivotCache cacheId="2" r:id="rId31"/><pivotCache cacheId="3" r:id="rId32"/></x15:pivotCaches>')
wb = wb.replace('<x15:pivotTableReference r:id="rId6"/></x15:pivotTableReferences>',
                '<x15:pivotTableReference r:id="rId6"/><x15:pivotTableReference r:id="rId33"/><x15:pivotTableReference r:id="rId34"/></x15:pivotTableReferences>')
assert 'rId31' in wb and 'rId33' in wb
put('xl/workbook.xml', wb)

ct = txt('[Content_Types].xml')
add = ''.join(
    f'<Override PartName="/xl/pivotCache/pivotCacheDefinition{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
    f'<Override PartName="/xl/pivotTables/pivotTable{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>'
    f'<Override PartName="/xl/charts/chart{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
    f'<Override PartName="/xl/charts/style{n}.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>'
    f'<Override PartName="/xl/charts/colors{n}.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>' for n in (3, 4))
put('[Content_Types].xml', ct.replace('</Types>', add + '</Types>'))

# ---------------------------------------------------------------- write (keep original part order, append new parts)
order = zin.namelist() + [n for n in parts if n not in zin.namelist()]
with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in order:
        info = zipfile.ZipInfo(n, date_time=(2026, 9, 27, 12, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, parts[n])
print('wrote', DST, 'manager order (asc):', bcats)
