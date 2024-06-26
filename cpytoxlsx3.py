"""Copy data from a physical or logical file to an Excel binary file.

Written by John Yeung.  Last modified 2024-04-12.

Usage (assuming Richard Schoen's QshOni is installed):
    qshoni/qshpyrun &script_lib 'cpytoxlsx3.py' (
        &pf
        &xlsx
        [&A1_text
        &A2_text
        ...])
        &py_version

Required parameters:
    &script_lib = IFS directory containing cpytoxlsx3.py
    &pf = qualified name of file to copy
    &xlsx = name of workbook to create, including path and extension
    &py_version = Python version; must be at least 3.6

Optional parameters:
    &A1_text = free-form text to appear at the top of the spreadsheet
    &A2_text = free-form text to appear on the 2nd line of the spreadsheet

    If using QshOni, the limit is 38 of these (because QSHPYRUN provides up
    to 40 total arguments to the Python script), with each limited to 200
    characters.

Dependencies:
  - Python 3.6 or later, installed via yum
  - PyODBC, installed via yum
  - XlsxWriter, installed via pip
  - QshOni (see https://github.com/richardschoen/qshoni)

yum can be invoked in a PASE shell or through ACS (go to the Tools menu,
choose Open Source Package Management).

QshOni is optional, but having it simplifies the Python call.

[To reduce confusion between Excel columns and database table columns,
DDS-based terminology will be used below unless otherwise noted.]

Features:

  - Column headings come from the field headings (as would be defined
    by the COLHDG keyword).  If there are multiple lines in a heading,
    they are trimmed and then joined into one line for the spreadsheet.
  - For any field without a heading, the field name is used instead.
    Field text is NOT checked for this purpose.  To specify a blank
    column heading (and avoid using the field name), define the field
    with COLHDG('*BLANK').
  - Column headings wrap and are displayed in bold.

  - Each column is sized approximately according to its longest data,
    assuming that the font is Calibri 11, unless a width is specified
    in the field text (using 'width=<number>').  [For this purpose,
    the length of numeric data is assumed to always include commas and
    fixed decimal places.]
  - Numeric data may be formatted using an Excel format string in the
    field text (using 'format="<string>"').
  - If there is no format string, but the field has a supported EDTCDE
    value, the column will be formatted according to the edit code.
    The supported edit codes are 1, 2, 3, 4, N, O, P, and Q.
  - For fields defined as character, the column will be set to Excel
    text format (to make it harder to accidentally convert digit-only
    character data into numeric by "visiting" the cell in Excel).
  - If 'zero=blank' is specified in the field text, cells which would
    have been zero are empty instead.  (This is different than using a
    format string or edit code to hide zero values.  See the ISBLANK
    and ISNUMBER functions in Excel.)
  - If 'wrap=on' is specified in the field text, the contents of the
    cell will wrap.  The row height is adjusted automatically to
    accommodate multiple lines.
  - Columns may be skipped entirely by specifying COLHDG('*SKIP').
  - Numeric fields that are 8 digits long with no decimal places are
    automatically converted to dates if they have a suitable edit word.
  - Numeric fields that are 6 digits long with no decimal places are
    automatically converted to times if they have a suitable edit word.
  - Free-form text inserted at the top using the optional parameters
    will be displayed in bold.
  - Blank rows may be inserted when the value in a particular field
    changes by specifying 'break on <fieldname>' in the record (not
    field!) text.  (I am not sure this can be done with SQL.)

The motivation for this program is to provide a tool for easy generation
of formatted spreadsheets.

Nice-to-have features not yet implemented include general edit word
support (not just for date detection), more available edit codes, more
comprehensive date support, the ability to choose fonts, and automatic
population of multiple sheets given multiple file members.

Also, it would be nice to wrap this in a command for even greater ease of
use, including meaningful promptability.
"""

import sys
import re
from os import system
from datetime import date, time, datetime
from decimal import Decimal

import pyodbc  # install with yum
from xlsxwriter.workbook import Workbook  # install with pip

# Table of empirically determined character widths in Calibri 11, the default
# font used by XlsxWriter.  Note that font rendering is somewhat dependent
# on the configuration of the PC that is used to open the resulting file, so
# these widths are not necessarily exact for other people's PCs.  Also note
# that only basic ASCII letters, digits, and punctuation are defined here;
# any other characters will be assumed to be the same size as a digit.
char_groups = {
    '0123456789agkvxyEFSTYZ#$*+<=>?^_|~': 203.01 / 28,
    'bdehnopquBCKPRX': 232.01 / 28,
    'cszL"/\\': 174.01 / 28,
    'frtJ!()-[]{}': 145.01 / 28,
    'ijlI,.:;`': 116.01 / 28,
    'mM': 348.01 / 28,
    'w%': 319.01 / 28,
    'ADGHUV': 261.01 / 28,
    'NOQ&': 290.01 / 28,
    'W@': 377.01 / 28,
    "' ": 87.01 / 28}
pixel_widths = {}
for group, width in char_groups.items():
    for char in group:
        pixel_widths[char] = width

bold_char_groups = {
    '0123456789agkvxyEFSTZ"#$*+<=>?^_|~': 203.01 / 28,
    'bdehnopquBCKPRXY': 232.01 / 28,
    'cszL/\\': 174.01 / 28,
    'frtJ!()-[]`{}': 145.01 / 28,
    "ijlI',.:;": 116.01 / 28,
    'm': 348.01 / 28,
    'w%&': 319.01 / 28,
    'ADHV': 261.01 / 28,
    'GNOQU': 290.01 / 28,
    'M@': 377.01 / 28,
    'W': 406.01 / 28,
    ' ': 87.01 / 28}
bold_pixel_widths = {}
for group, width in bold_char_groups.items():
    for char in group:
        bold_pixel_widths[char] = width

class SheetInfo(object):
    def __init__(self):
        self.fieldlist = []  # fields to include, not necessarily all
        self.headings = {}
        self.numformats = {}
        self.dateflags = {}
        self.timeflags = {}
        self.commaflags = {}
        self.decplaces = {}
        self.colwidths = {}
        self.wrapped = {}
        self.blankzeros = {}
        self.breakfield = None

##
##  Miscellaneous utilities
##

def quoted(s):
    return "'" + s.replace("'", "''") + "'"

def qcmdexc(cs, cmd):
    cs.execute(f"values qcmdexc({quoted(cmd)})")
    result = cs.fetchone()[0]  # 1 for success, -1 for failure
    if result != 1:
        raise RuntimeError(
            f"Error while trying to execute the following command:\n{cmd}")

def fetch(cs):
    """Do `fetchall` and post-process the result.

    - Right-trim all strings.
    - Replace decimals (packed or zoned) with integers if scale is zero.
    - Return a list of tuples instead of a list of Row objects.
    """
    result = cs.fetchall()
    if result:
        for row in result:
            for cx, col in enumerate(cs.description):
                if row[cx] is None:
                    continue
                if col[1] == str:
                    row[cx] = row[cx].rstrip()
                elif col[1] == Decimal and col[5] == 0:
                    row[cx] = int(row[cx])
        return [tuple(row) for row in result]
    return result

# TO DO: Flesh this out into a proper SNDMSG wrapper.  It should be possible
# to determine the current user profile with some SQL view or service.
def sndmsg(msg):
    print(msg)

##
##  Width calculations
##

def integer_digits(n):
    """Return the number of digits in a positive integer."""
    if n == 0:
        return 1
    digits = 0
    while n:
        digits += 1
        n //= 10
    return digits

def number_analysis(n, dp=0):
    """Return a 4-tuple of (digits, thousands, points, signs)."""
    digits, thousands, points, signs = 0, 0, 0, 0
    if n < 0:
        signs = 1
        n = -n
    if dp > 0:
        points = 1
    idigits = integer_digits(int(n))
    digits = idigits + points + dp
    thousands = (idigits - 1) // 3
    return digits, thousands, points, signs

def colwidth_from_pixels(pixels):
    """Convert pixels to the user-facing units presented by Excel."""
    # Excel has a mysterious fudge factor when autofitting
    if pixels > 34:
        pixels -= 1
    if pixels > 62:
        pixels -= 1
    # The first unit of column width is 12 pixels; each subsequent unit is 7
    if pixels < 12:
        return pixels / 12.0
    return (pixels - 5) / 7.0

def textwidth(data, bold=False):
    """Try to autofit text data."""
    charwidths = bold_pixel_widths if bold else pixel_widths
    pixels = 7
    for char in str(data):
        if char in charwidths:
            pixels += charwidths[char]
        else:
            pixels += charwidths['0']
    return colwidth_from_pixels(pixels)

def numwidth(data, dp, use_commas=False):
    """Try to autofit a number.

    Note that in Calibri 11, characters used in numbers do not change
    width when bold.
    """
    charwidths = pixel_widths
    digits, commas, points, signs = number_analysis(data, dp)
    pixels = 7 + digits * charwidths['0']
    if use_commas:
        pixels += commas * charwidths[',']
    pixels += points * charwidths['.']
    pixels += signs * charwidths['-']
    return colwidth_from_pixels(pixels)

def datewidth():
    """Set aside enough width for an 8-character date, with separators."""
    charwidths = pixel_widths
    pixels = 7 + 8 * charwidths['0'] + 2 * charwidths['/']
    return colwidth_from_pixels(pixels)

def timewidth(bold=False):
    """Set aside enough width for an HH:MM:SS AM/PM time."""
    charwidths = bold_pixel_widths if bold else pixel_widths
    digits_width = 6 * charwidths['0']
    sep_width = 2 * charwidths[':']
    space_width = charwidths[' ']
    am_pm_width = max(charwidths['A'], charwidths['P']) + charwidths['M']
    pixels = 7 + digits_width + sep_width + space_width + am_pm_width
    return colwidth_from_pixels(pixels)

def timestampwidth(bold=False):
    """Set aside enough width for an MM/DD/YY HH:MM:SS AM/PM timestamp."""
    charwidths = bold_pixel_widths if bold else pixel_widths
    digits_width = 12 * charwidths['0']
    sep_width = 2 * charwidths['/'] + 2 * charwidths[':']
    space_width = 2 * charwidths[' ']
    am_pm_width = max(charwidths['A'], charwidths['P']) + charwidths['M']
    pixels = 7 + digits_width + sep_width + space_width + am_pm_width
    return colwidth_from_pixels(pixels)

##
##  Number formatting
##

def default_numformat(dp=0, use_commas=False):
    """Generate a style dictionary for Excel fixed number format."""
    integers, decimals = '0', ''
    if use_commas:
        integers = '#,##0'
    if dp > 0:
        decimals = '.' + '0' * dp
    return {'num_format': integers + decimals}

def editcode(code, dp=0):
    """Generate a style dictionary corresponding to an edit code."""
    code = code.lower()
    if len(code) != 1 or code not in ('1234nopq'):
        return default_numformat(dp)
    sign, integers, decimals, zero = '', '#', '', ''
    if code in 'nopq':
        sign = '-'
    if code in '12no':
        integers = '#,###'
    if dp > 0:
        decimals = '.' + '0' * dp
    positive = integers + decimals
    negative = sign + positive
    if code in '13np':
        zero = positive[:-1] + '0'
    return {'num_format': ';'.join((positive, negative, zero))}

def is_numeric_date(size, editword):
    return size == (8, 0) and editword in ("'    -  -  '", "'    /  /  '")

def is_numeric_time(size, editword):
    return size == (6, 0) and editword in ("'  .  .  '", "'  :  :  '")

##
##  Process DDS
##

def sheetinfo(cs, libname, filename):
    """Retrieve formatting information from DSPFFD.

    It would be nice to use the SQL facilities that are now available,
    such as QSYS2.SYSCOLUMNS2, but so far I have not been able to find
    the record text anywhere but DSPFFD.
    """
    qcmdexc(cs,
        f"dspffd {libname}/{filename}"
        ' output(*outfile) outfile(qtemp/dspffdpf)')
    cs.execute('''
        select
            whflde, whftxt, whchd1, whchd2, whchd3, whtext,
            whfldd, whfldp, whfldb,
            whecde, whewrd
        from qtemp.dspffdpf''')
    all_fields = [t[0] for t in cs.description]
    s = SheetInfo()
    s.fieldlist = []  # fields to include in spreadsheet
    s.headings = {}
    s.numformats = {}
    s.dateflags = {}
    s.timeflags = {}
    s.commaflags = {}
    s.decplaces = {}
    s.colwidths = {}
    s.wrapped = {}
    s.blankzeros = {}
    s.breakfield = None

    for row in fetch(cs):
        fieldname = row[0]  # column_name in SYSCOLUMNS2
        fieldtext = row[1]  # column_text in SYSCOLUMNS2
        colhdg = row[2:5]  # column_heading in SYSCOLUMNS2
        rcdtext = row[5]  # don't know where to find this other than DSPFFD
        digits, decimal_places, byte_length = row[6:9]
        edtcde, edtwrd = row[9:11]

        # Set break field.
        if s.breakfield is None:
            match = re.search(r'break on (\S+)', rcdtext, re.IGNORECASE)
            if match:
                s.breakfield = match.group(1).upper()
            if s.breakfield not in all_fields:
                s.breakfield = ''  # different than None; prevent recalculation

        # Set Excel column heading from DDS headings.  If those are blank,
        # use the field name instead.  There are special values to exclude
        # the field entirely or set the Excel column heading to blank.
        heading = ' '.join(colhdg).strip()
        if heading.upper() == '*SKIP':
            continue
        if heading.upper() in ('*BLANK', '*BLANKS'):
            heading = ''
        elif not heading:
            heading = ''
        s.fieldlist.append(fieldname)
        s.headings[fieldname] = heading

        # Set field size and type.
        if digits:
            fieldsize = (digits, decimal_places)
            s.decplaces[fieldname] = fieldsize[1]
            numeric = True
        else:
            fieldsize = byte_length
            numeric = False

        # Look for number format string.
        match = re.search(r'format="(.*)"', fieldtext, re.IGNORECASE)
        if match:
            numformat = {'num_format': match.group(1)}
        elif numeric:
            numformat = editcode(edtcde, decimal_places)
        else:
            numformat = None
        if numformat:
            s.numformats[fieldname] = numformat
            s.commaflags[fieldname] = ',' in numformat['num_format']

        # Check whether it looks like a numeric date or time.
        s.dateflags[fieldname] = is_numeric_date(fieldsize, edtwrd)
        s.timeflags[fieldname] = is_numeric_time(fieldsize, edtwrd)

        # Look for fixed column width.
        match = re.search(r'width=([1-9][0-9]*)', fieldtext, re.IGNORECASE)
        if match:
            s.colwidths[fieldname] = int(match.group(1))

        # Look for text wrap flag.
        match = re.search(r'wrap=(\*)?on', fieldtext, re.IGNORECASE)
        if match:
            s.wrapped[fieldname] = True

        # Look for zero-suppression flag.
        match = re.search(r'zero(s|es)?=blanks?', fieldtext, re.IGNORECASE)
        if match:
            s.blankzeros[fieldname] = True

    return s

##
##  Main logic
##

def main():
    # Check parameters.
    parameters = len(sys.argv) - 1
    if parameters < 2:
        sndmsg(f"Program needs at least 2 parameters; received {parameters}.")
        sys.exit(2)
    pf = sys.argv[1].split('/')
    if len(pf) == 1:
        libname = '*LIBL'
        filename = pf[0].upper()
    elif len(pf) == 2:
        libname = pf[0].upper()
        filename = pf[1].upper()
    else:
        sndmsg('Could not parse file name.')
        sys.exit(2)
    sndmsg('Parameters checked.')

    # Connect to the database.
    conn = pyodbc.connect(dsn='*LOCAL')
    c1 = conn.cursor()

    # Get column headings and formatting information from the DDS.
    fmt = sheetinfo(c1, libname, filename)

    # Create a workbook with one sheet
    with Workbook(sys.argv[2]) as wb:
        ws = wb.add_worksheet(filename)
        rx = 0

        title_style = wb.add_format({'bold': True})
        header_style = wb.add_format({'bold': True, 'text_wrap': True})
        date_style = wb.add_format({'num_format': 'm/d/yyyy'})
        time_style = wb.add_format({'num_format': 'h:mm:ss AM/PM'})
        timestamp_style = wb.add_format({'num_format': 'm/d/yy h:mm:ss AM/PM'})
        text_style = wb.add_format({'num_format': '@'})
        wrapped_style = wb.add_format({'text_wrap': True})
        for field, format_dict in fmt.numformats.items():
            fmt.numformats[field] = wb.add_format(format_dict)

        # Populate first few rows using additional parameters, if provided.
        # Typically, these rows would be used for report ID, date, and title.
        if parameters > 2:
            for arg in sys.argv[3:]:
                ws.write_string(rx, 0, arg, title_style)
                rx += 1
            rx += 1  # skip a row before starting the column headings

        # Keep track of the widest data in each column.
        maxwidths = [0] * len(fmt.fieldlist)

        # Create a row for column headings.
        for col, name in enumerate(fmt.fieldlist):
            desc = fmt.headings[name]
            ws.write_string(rx, col, desc, header_style)
            if name not in fmt.colwidths:
                maxwidths[col] = textwidth(desc, bold=True)

        bx = None  # ordinal position of break field
        breakvalue = None
        if fmt.breakfield:
            bx = fmt.fieldlist.index(fmt.breakfield)

        # Read from database and write to spreadsheet, row by row.
        c1.execute(f"select * from {libname}.{filename}")
        sndmsg('Opened ' + libname + '/' + filename + ' for reading.')

        for row in fetch(c1):
            rx += 1
            if bx is not None:
                if row[bx] != breakvalue and breakvalue is not None:
                    rx += 1
                breakvalue = row[bx]
            for cx, value in enumerate(row):
                fieldname = fmt.fieldlist[cx]
                nativedate = False
                nativetime = False
                nativetimestamp = False
                # When using `isinstance`, class `datetime` has to be checked
                # before `date` because `datetime` is a subclass of `date`.
                if isinstance(value, datetime):
                    ws.write_datetime(rx, cx, value, timestamp_style)
                    nativetimestamp = True
                elif isinstance(value, date):
                    ws.write_datetime(rx, cx, value, date_style)
                    nativedate = True
                elif isinstance(value, time):
                    ws.write_datetime(rx, cx, value, time_style)
                    nativetime = True
                elif fmt.dateflags[fieldname]:
                    if value:
                        year, md = divmod(value, 10000)
                        month, day = divmod(md, 100)
                        ws.write_datetime(
                            rx, cx, date(year, month, day), date_style)
                elif fmt.timeflags[fieldname]:
                    if value:
                        hour, minsec = divmod(value, 10000)
                        minute, second = divmod(minsec, 100)
                        ws.write_datetime(
                            rx, cx, time(hour, minute, second), time_style)
                elif value == 0 and fieldname in fmt.blankzeros:
                    pass
                elif fieldname in fmt.numformats:
                    ws.write(rx, cx, value, fmt.numformats[fieldname])
                elif fieldname in fmt.wrapped:
                    ws.write(rx, cx, value, wrapped_style)
                elif isinstance(value, str):
                    ws.write_string(rx, cx, value, text_style)
                else:
                    ws.write(rx, cx, value)
                if fieldname not in fmt.colwidths:
                    if nativedate or fmt.dateflags[fieldname]:
                        maxwidths[cx] = datewidth()
                    elif nativetime or fmt.timeflags[fieldname]:
                        maxwidths[cx] = timewidth()
                    elif nativetimestamp:
                        maxwidths[cx] = timestampwidth()
                    if fieldname in fmt.decplaces:
                        dp = fmt.decplaces[fieldname]
                        cf = fmt.commaflags[fieldname]
                        maxwidths[cx] = max(maxwidths[cx], numwidth(value, dp, cf))
                    else:
                        maxwidths[cx] = max(maxwidths[cx], textwidth(value))

        # Set column widths
        for cx in range(len(fmt.fieldlist)):
            if fmt.fieldlist[cx] in fmt.colwidths:
                ws.set_column(cx, cx, fmt.colwidths[fmt.fieldlist[cx]])
            else:
                ws.set_column(cx, cx, maxwidths[cx])

        sndmsg('File copied to ' + sys.argv[2] + '.')

##
##  Script entry point
##

if __name__ == '__main__':
    main()
