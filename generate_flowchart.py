#!/usr/bin/env python3
"""
generate_flowchart.py
Generates an Excel flow chart (.xlsm) from a list of production phases,
including the generation date under the directional arrow.
"""

import zipfile
import shutil
import os
import uuid
import argparse
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURE PHASES HERE (used when executing the script directly)
# ─────────────────────────────────────────────────────────────────────────────
PHASES_DATA = [
    {"number": 10, "phase": "Raw Material", "ext_int": "External", "department": "Supplier"},
    {"number": 20, "phase": "Heat Treatment", "ext_int": "Internal", "department": "HT Dpt."},
    {"number": 30, "phase": "Grinding", "ext_int": "Internal", "department": "Grinding Dpt.",
     "note": "Pay attention to diameters"},
    {"number": 40, "phase": "Surface Treatment", "ext_int": "External", "department": "Supplier"},
    {"number": 50, "phase": "Assembly", "ext_int": "Internal", "department": "Assembly"},
    {"number": 60, "phase": "Final Testing", "ext_int": "Internal", "department": "Quality",
     "note": "100% visual inspection"},
    {"number": 70, "phase": "Shipping", "ext_int": "Internal", "department": "Logistics"},
]

PRODUCT_TITLE = "PRODUCT NAME"


# ─────────────────────────────────────────────────────────────────────────────
# ── XML helpers ──────────────────────────────────────────────────────────────

def _new_id():
    return str(uuid.uuid4()).upper()


def _roundrect_shape(shape_id, name, col_from, col_to, row_from, row_to,
                     x_off=0, col_off=0, row_off_to=1):
    return f"""<xdr:twoCellAnchor>
<xdr:from><xdr:col>{col_from}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>{col_to}</xdr:col><xdr:colOff>{col_off}</xdr:colOff><xdr:row>{row_to}</xdr:row><xdr:rowOff>{row_off_to}</xdr:rowOff></xdr:to>
<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="{name}"><a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{{{_new_id()}}}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>
<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>
<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/><a:lstStyle/><a:p><a:pPr algn="l"/><a:endParaRPr lang="en-GB" sz="1100"/></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"""


def _arrow_shape(n_phases):
    row_end = 5 + 3 * n_phases
    # FINAL SOLUTION: wrap="square" + normAutofit + non-breaking space (&#160;) in text
    return f"""<xdr:twoCellAnchor>
<xdr:from><xdr:col>8</xdr:col><xdr:colOff>446314</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>16329</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>12</xdr:col><xdr:colOff>234043</xdr:colOff><xdr:row>{row_end}</xdr:row><xdr:rowOff>13607</xdr:rowOff></xdr:to>
<xdr:sp macro="[0]!DynamicArrow" textlink=""><xdr:nvSpPr><xdr:cNvPr id="25" name="DynamicArrow" descr="Production Flow"><a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{{{_new_id()}}}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr><a:spLocks/></xdr:cNvSpPr></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="9984921" y="2628900"/><a:ext cx="2128158" cy="7181850"/></a:xfrm><a:prstGeom prst="downArrow"><a:avLst><a:gd name="adj1" fmla="val 50000"/><a:gd name="adj2" fmla="val 41593"/></a:avLst></a:prstGeom><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:ln><a:solidFill><a:schemeClr val="bg2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln></xdr:spPr>
<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>
<xdr:txBody><a:bodyPr vert="vert270" wrap="square" vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-GB" sz="3200" b="1"><a:latin typeface="Times New Roman"/><a:cs typeface="Times New Roman"/></a:rPr><a:t>PRODUCTION&#160;FLOW</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"""


def _logo_picture():
    return """<xdr:twoCellAnchor editAs="oneCell">
<xdr:from><xdr:col>1</xdr:col><xdr:colOff>54702</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>173128</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>1</xdr:col><xdr:colOff>1164771</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>292237</xdr:rowOff></xdr:to>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="108" name="Picture 11" descr="Logo"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{21B8031A-36D7-46D8-BAC5-18A331A58741}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr bwMode="auto"><a:xfrm><a:off x="664302" y="173128"/><a:ext cx="1110069" cy="772252"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525"><a:noFill/><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"""


def _title_box(product_title):
    return f"""<xdr:twoCellAnchor>
<xdr:from><xdr:col>1</xdr:col><xdr:colOff>1250769</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>239487</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>8</xdr:col><xdr:colOff>598714</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>284988</xdr:rowOff></xdr:to>
<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="110" name="Rounded rectangle 27"><a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{{{_new_id()}}}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="1860369" y="239487"/><a:ext cx="8600802" cy="698644"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="bg2"><a:lumMod val="90000"/></a:schemeClr></a:solidFill><a:bevel/></a:ln></xdr:spPr>
<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>
<xdr:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"><a:defRPr/></a:pPr><a:r><a:rPr lang="en-GB" sz="1400" b="1"><a:latin typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/><a:cs typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/></a:rPr><a:t>PRODUCTION FLOW CHART</a:t></a:r></a:p>
<a:p><a:pPr algn="ctr"><a:defRPr/></a:pPr><a:r><a:rPr lang="en-GB" sz="2000" b="1"><a:latin typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/><a:cs typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/></a:rPr><a:t>{product_title}</a:t></a:r></a:p>
</xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"""


def _phases_banner():
    return """<xdr:twoCellAnchor>
<xdr:from><xdr:col>1</xdr:col><xdr:colOff>10885</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>423562</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>8</xdr:col><xdr:colOff>587828</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>348344</xdr:rowOff></xdr:to>
<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="109" name="Rounded rectangle 26"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{3A41C090-7145-47FF-9440-C8C121CF7DBF}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="620485" y="1076705"/><a:ext cx="9829800" cy="1231068"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="bg2"/></a:solidFill></a:ln></xdr:spPr>
<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>
<xdr:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"><a:defRPr/></a:pPr><a:r><a:rPr lang="en-GB" sz="3200" b="1"><a:latin typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/><a:cs typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/></a:rPr><a:t>PRODUCTION PHASES</a:t></a:r></a:p>
</xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"""


def _notes_box(notes_list, start_row):
    """Generates the 'Additional remarks:' graphic box at the bottom of the Flow Chart."""
    paragraphs = [
        '<a:p><a:pPr algn="l"><a:defRPr/></a:pPr><a:r><a:rPr lang="en-GB" sz="2000" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Times New Roman"/><a:cs typeface="Times New Roman"/></a:rPr><a:t>Additional remarks:</a:t></a:r></a:p>'
    ]
    for note in notes_list:
        safe_note = _esc(note)
        paragraphs.append(
            f'<a:p><a:pPr algn="l"><a:defRPr/></a:pPr><a:r><a:rPr lang="en-GB" sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Times New Roman"/><a:cs typeface="Times New Roman"/></a:rPr><a:t>{safe_note}</a:t></a:r></a:p>'
        )

    paragraphs_xml = '\n'.join(paragraphs)
    end_row = start_row + len(notes_list) + 1
    shape_id = 1000 + start_row

    return f"""<xdr:twoCellAnchor>
<xdr:from><xdr:col>1</xdr:col><xdr:colOff>10000</xdr:colOff><xdr:row>{start_row}</xdr:row><xdr:rowOff>100000</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{end_row}</xdr:row><xdr:rowOff>100000</xdr:rowOff></xdr:to>
<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="Notes Box"><a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{{{_new_id()}}}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill><a:bevel/></a:ln></xdr:spPr>
<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style>
<xdr:txBody><a:bodyPr wrap="square" anchor="t" lIns="91440" tIns="91440" rIns="91440" bIns="91440"/><a:lstStyle/>
{paragraphs_xml}
</xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"""


def _date_box(date_str, start_row):
    """Generates a transparent text box with the date under the arrowhead."""
    end_row = start_row + 1
    shape_id = 2000 + start_row
    return f"""<xdr:twoCellAnchor>
<xdr:from><xdr:col>8</xdr:col><xdr:colOff>446314</xdr:colOff><xdr:row>{start_row}</xdr:row><xdr:rowOff>150000</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>12</xdr:col><xdr:colOff>234043</xdr:colOff><xdr:row>{end_row}</xdr:row><xdr:rowOff>150000</xdr:rowOff></xdr:to>
<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="Update Date"><a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{{{_new_id()}}}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>
<xdr:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style>
<xdr:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"><a:defRPr/></a:pPr><a:r><a:rPr lang="en-GB" sz="1800" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Times New Roman"/><a:cs typeface="Times New Roman"/></a:rPr><a:t>Date: {date_str}</a:t></a:r></a:p>
</xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"""


# ── File Construction ─────────────────────────────────────────────────────────

def build_drawing_xml(phases, product_title):
    ns = (
        'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    )
    parts = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr {ns}>']

    parts.append(_logo_picture())
    parts.append(_title_box(product_title))
    parts.append(_phases_banner())

    parts.append(_roundrect_shape(36, "Header_OpNum", 1, 2, 4, 5, col_off=1, row_off_to=1))
    parts.append(_roundrect_shape(37, "Header_Phase", 3, 4, 4, 5, col_off=8467, row_off_to=1))
    parts.append(_roundrect_shape(38, "Header_ExtInt", 5, 6, 4, 5, col_off=8467, row_off_to=1))
    parts.append(_roundrect_shape(44, "Header_Dept", 7, 8, 4, 5, col_off=8467, row_off_to=1))

    parts.append(_arrow_shape(len(phases)))

    # Shapes for each phase
    sid = 100
    notes_list = []

    for i, phase in enumerate(phases, start=1):
        draw_row = 4 + 3 * i
        parts.append(_roundrect_shape(sid, f"Op{i}_Num", 1, 2, draw_row, draw_row + 1, col_off=1, row_off_to=1))
        parts.append(
            _roundrect_shape(sid + 1, f"Op{i}_Phase", 3, 4, draw_row, draw_row + 1, col_off=8467, row_off_to=1))
        parts.append(
            _roundrect_shape(sid + 2, f"Op{i}_ExtInt", 5, 6, draw_row, draw_row + 1, col_off=8467, row_off_to=1))
        parts.append(_roundrect_shape(sid + 3, f"Op{i}_Dept", 7, 8, draw_row, draw_row + 1, col_off=8467, row_off_to=1))
        sid += 4

        # Intercept and format notes
        note_text = phase.get("note", "").strip()
        if note_text:
            notes_list.append(f"* Op. {phase.get('number')} ({phase.get('phase')}): {note_text}")

    # --- DATE AND NOTES ---

    # The arrow ends exactly at this row
    row_end_arrow = 5 + 3 * len(phases)

    # Inserts today's date positioned at the end of the arrow
    current_date = datetime.now().strftime("%d/%m/%Y")
    parts.append(_date_box(current_date, row_end_arrow))

    # Adds the Notes box at the end (moved slightly lower to avoid overlapping)
    if notes_list:
        row_start_notes = row_end_arrow + 2
        parts.append(_notes_box(notes_list, row_start_notes))

    parts.append('</xdr:wsDr>')
    return '\n'.join(parts)


def build_sheet_xml(phases):
    n = len(phases)
    last_data_row = 5 + 3 * n
    last_row = last_data_row + 1

    rows = []

    rows.append(
        '<row r="5" spans="1:41" s="3" customFormat="1" ht="52.2" customHeight="1">'
        '<c r="A5" s="10"/>'
        '<c r="B5" s="2" t="inlineStr"><is><t>Number of Operation</t></is></c>'
        '<c r="C5" s="10"/>'
        '<c r="D5" s="2" t="inlineStr"><is><t>Production Phases</t></is></c>'
        '<c r="E5" s="10"/>'
        '<c r="F5" s="2" t="inlineStr"><is><t>External/ Internal</t></is></c>'
        '<c r="G5" s="10"/>'
        '<c r="H5" s="2" t="inlineStr"><is><t>Department</t></is></c>'
        '</row>'
    )

    rows.append('<row r="6" spans="1:41" ht="26" customHeight="1"/>')
    rows.append('<row r="7" spans="1:41" ht="26" customHeight="1"/>')

    for i, phase in enumerate(phases, start=1):
        data_row = 5 + 3 * i
        conn_row1 = data_row + 1
        conn_row2 = data_row + 2

        num = phase.get("number", i * 10)
        ph = phase.get("phase", "")
        ei = phase.get("ext_int", "")
        dep = phase.get("department", "")

        # Magically adds the asterisk if the phase has a note
        if phase.get("note", "").strip():
            ph += " *"

        rows.append(
            f'<row r="{data_row}" spans="1:41" ht="52.2" customHeight="1">'
            f'<c r="B{data_row}" s="2"><v>{num}</v></c>'
            f'<c r="D{data_row}" s="2" t="inlineStr"><is><t>{_esc(ph)}</t></is></c>'
            f'<c r="F{data_row}" s="2" t="inlineStr"><is><t>{_esc(ei)}</t></is></c>'
            f'<c r="H{data_row}" s="2" t="inlineStr"><is><t>{_esc(dep)}</t></is></c>'
            f'</row>'
        )

        rows.append(f'<row r="{conn_row1}" spans="1:41" ht="26" customHeight="1"/>')
        rows.append(f'<row r="{conn_row2}" spans="1:41" ht="26" customHeight="1"/>')

    rows_xml = '\n'.join(rows)
    dim_ref = f"A5:H{last_data_row}"

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
           mc:Ignorable="x14ac">
  <sheetPr codeName="Sheet1"><tabColor theme="0"/></sheetPr>
  <dimension ref="{dim_ref}"/>
  <sheetViews>
    <sheetView showGridLines="0" tabSelected="1" zoomScale="70" zoomScaleNormal="70" workbookViewId="0">
      <selection activeCell="P1" sqref="P1"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultColWidth="8.88671875" defaultRowHeight="52.2" customHeight="1"/>
  <cols>
    <col min="1" max="1" width="8.88671875" style="11"/>
    <col min="2" max="2" width="24.6640625" style="10" customWidth="1"/>
    <col min="3" max="3" width="8.88671875" style="11"/>
    <col min="4" max="4" width="31.88671875" style="11" customWidth="1"/>
    <col min="5" max="5" width="8.88671875" style="11"/>
    <col min="6" max="6" width="30.33203125" style="11" customWidth="1"/>
    <col min="7" max="7" width="8.88671875" style="11"/>
    <col min="8" max="8" width="21.33203125" style="11" customWidth="1"/>
    <col min="9" max="16384" width="8.88671875" style="1"/>
  </cols>
  <sheetData>
{rows_xml}
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="portrait" r:id="rId1"/>
  <drawing r:id="rId2"/>
</worksheet>"""


def _esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


# ── Entry point ───────────────────────────────────────────────────────────────

def generate(template_path, output_path, phases, title=None):
    product_title = title or PRODUCT_TITLE

    shutil.copy2(template_path, output_path)

    drawing_xml = build_drawing_xml(phases, product_title)
    sheet_xml = build_sheet_xml(phases)

    tmp = output_path + ".tmp.zip"
    with zipfile.ZipFile(output_path, 'r') as zin, \
            zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            if item.filename == 'xl/drawings/drawing1.xml':
                zout.writestr(item, drawing_xml.encode('utf-8'))
            elif item.filename == 'xl/worksheets/sheet1.xml':
                zout.writestr(item, sheet_xml.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))

    os.replace(tmp, output_path)
    print(f"✅ Flow Chart generated: {output_path} ({len(phases)} phases)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generates Excel flow chart from production phases")
    parser.add_argument("--template", default="TEMPLATE_FLOW_CHART.xlsm", help=".xlsm template path")
    parser.add_argument("--output", default="flow_chart_output.xlsm", help="Output file")
    parser.add_argument("--title", default=None, help="Flow chart subtitle")
    args = parser.parse_args()

    # Test with a few phases to verify autofit
    generate(template_path=args.template, output_path=args.output, phases=PHASES_DATA[:2], title=args.title)
