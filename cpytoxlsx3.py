### UNDER CONSTRUCTION

# Trying to convert to Python 3, running in PASE. This requires pyodbc.
# There was some way to connect to the local IBM i without any setup or
# credentials. I don't remember at the moment, but let's assume I can
# find that out easily enough.
#
# Overview: The main thing is converting the RLA to SQL. Also necessary,
# but minor, is converting Python 2 to 3. Finally, make use of the SQL
# views instead of doing DSPFFD to an output file.

"""Copy data from a physical or logical file to an Excel binary file.

Written by John Yeung.  Last modified 2015-04-29.

Usage (from CL):
    python27/python '/util/cpytoxlsx.py' parm(&pf &xlsx [&A1text &A2text ...])

The above assumes this program is located in '/util', and that iSeriesPython
2.7 is installed.  However, you can put this program anywhere you like in
the IFS.  You can also probably use iSeriesPython 2.5, but this is not tested
and not recommended.  The XlsxWriter package is required.  Instructions for
downloading and installing it can be found at

http://iseriespython.blogspot.ca/2013/06/installing-python-packages-like.html

Some features/caveats:

-  Column headings come from the COLHDG values in the DDS.  Multiple
    values for a single field are joined by spaces, not newlines.  For
    any fields without a COLHDG, or with only blanks in the COLHDG (these
    two situations are indistinguishable), the field name is used as the
    heading (the TEXT keyword is not checked).  To specify a blank column
    heading rather than the field name, use COLHDG('*BLANK').
-  Column headings wrap and are displayed in bold.
-  Each column is sized approximately according to its longest data,
    assuming that the default font is Calibri 11, unless a width is
    specified in the field text (using 'width=<number>').  [For this
    purpose, the length of numeric data is assumed to always include
    commas and fixed decimal places.]
-  Each column may be formatted using an Excel format string in the
    field text (using 'format="<string>"').
-  Columns with a supported EDTCDE value but no format string are
    formatted according to the edit code.
-  Character fields with no format string and no edit code are set to
    Excel text format.
-  Columns may specify 'zero=blank' anywhere in the field text to leave
    a cell empty when its value is zero.  (This is different than using
    a format string or edit code to hide zero values.  See the ISBLANK
    and ISNUMBER functions in Excel.)
-  Columns may specify 'wrap=on' anywhere in the field text to wrap
    the contents.  This will automatically adjust the row height to
    accommodate multiple lines of text within the cell.
-  Columns may be skipped entirely by specifying COLHDG('*SKIP')
-  Numeric fields that are 8 digits long with no decimal places are
    automatically converted to dates if they have a suitable edit word.
-  Numeric fields that are 6 digits long with no decimal places are
    automatically converted to times if they have a suitable edit word.
-  Free-form data may be inserted at the top using additional parameters,
    one parameter for each row.  The data will be in bold.  Up to 13 of
    these additional parameters may be specified (because iSeriesPython
    accepts at most 15 parameters).
-  Blank rows may be inserted when the value in a particular field changes
    by specifying 'break on <fieldname>' in the record (not field!) text.

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
from datetime import date, time

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

# TO DO: Flesh this out into a proper SNDMSG wrapper. It should be possible
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
    if isinstance(n, float):
        idigits = integer_digits(int(n) + 1)
    elif isinstance(n, int):
        idigits = integer_digits(n)
    else:
        return None
    digits = idigits + dp
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


    # Have to figure out something to handle *LIBL. For now, require explicit
    # library.
    c1 = 'Some cursor'
    c1.execute(f"select * from {libname}.{filename}")
    sndmsg('Opened ' + libname + '/' + filename + ' for reading.')


    # Get column headings and formatting information from the DDS
    qcmdexc(
        f"dspffd {libname}/{filename}"
        ' output(*outfile) outfile(qtemp/dspffdpf)')
    c1.execute('''
        select whflde, whftxt, whchd1, whchd2, whchd3, whtext
        from qtemp.dspffdpf''')

    fieldlist = [t[0] for t in c1.description]
    headings = {}
    numformats = {}
    dateflags = {}
    timeflags = {}
    commaflags = {}
    decplaces = {}
    colwidths = {}
    wrapped = {}
    blankzeros = {}
    breakfield = None
    cmd = f"dspffd {libname}/{filename} output(*outfile) outfile(qtemp/dspffdpf)"
# Use QCMDEXC to execute the command.
c2 = 'Some other cursor'
c2.execute('select * from qtemp.dspffdpf')
ddsfile = c2.fetch()  # assumes `fetch` is a nice wrapper for `fetchall`

ddsfile.posf()
while not ddsfile.readn():
    fieldname = ddsfile['WHFLDE']  # column_name
    fieldtext = ddsfile['WHFTXT']  # column_text
    rcdtext = ddsfile['WHTEXT']  # don't know where to find this other than DSPFFD
    if breakfield is None:
        match = re.search(r'break on (\S+)', rcdtext, re.IGNORECASE)
        if match:
            breakfield = match.group(1).upper()
        if breakfield not in infile.fieldList():
            breakfield = ''  # different than None; prevent recalculation

    # Set heading
    headertuple = (ddsfile['WHCHD1'], ddsfile['WHCHD2'], ddsfile['WHCHD3'])
    text = ' '.join(headertuple).strip()
    if not text:
        text = fieldname
    elif text.upper() in ('*BLANK', '*BLANKS'):
        text = ''
    elif text.upper() == '*SKIP':
        continue
    fieldlist.append(fieldname)
    headings[fieldname] = text

    # Get field size and type
    if ddsfile['WHFLDD']:
        fieldsize = (ddsfile['WHFLDD'], ddsfile['WHFLDP'])
        decplaces[fieldname] = fieldsize[1]
        numeric = True
    else:
        fieldsize = ddsfile['WHFLDB']
        numeric = False

    # Look for number format string
    match = re.search(r'format="(.*)"', fieldtext, re.IGNORECASE)
    if match:
        numformat = {'num_format': match.group(1)}
    elif numeric:
        numformat = editcode(ddsfile['WHECDE'], ddsfile['WHFLDP'])
    else:
        numformat = None
    if numformat:
        numformats[fieldname] = numformat
        commaflags[fieldname] = ',' in numformat['num_format']

    # Check whether it looks like a numeric date or time
    dateflags[fieldname] = is_numeric_date(fieldsize, ddsfile['WHEWRD'])
    timeflags[fieldname] = is_numeric_time(fieldsize, ddsfile['WHEWRD'])

    # Look for fixed column width
    match = re.search(r'width=([1-9][0-9]*)', fieldtext, re.IGNORECASE)
    if match:
        colwidths[fieldname] = int(match.group(1))

    # Look for text wrap flag
    match = re.search(r'wrap=(\*)?on', fieldtext, re.IGNORECASE)
    if match:
        wrapped[fieldname] = True

    # Look for zero-suppression flag
    match = re.search(r'zero(s|es)?=blanks?', fieldtext, re.IGNORECASE)
    if match:
        blankzeros[fieldname] = True

ddsfile.close()

# Create a workbook with one sheet
wb = Workbook(sys.argv[2])
ws = wb.add_worksheet(infile.fileName())
row = 0

title_style = wb.add_format({'bold': True})
header_style = wb.add_format({'bold': True, 'text_wrap': True})
date_style = wb.add_format({'num_format': 'm/d/yyyy'})
time_style = wb.add_format({'num_format': 'h:mm:ss AM/PM'})
text_style = wb.add_format({'num_format': '@'})
wrapped_style = wb.add_format({'text_wrap': True})
for field, format_dict in numformats.items():
    numformats[field] = wb.add_format(format_dict)

# Populate first few rows using additional parameters, if provided.
# Typically, these rows would be used for report ID, date, and title.
if parameters > 2:
    for arg in sys.argv[3:]:
        ws.write_string(row, 0, arg, title_style)
        row += 1
    row += 1  # skip a row before starting the column headings

# Keep track of the widest data in each column
maxwidths = [0] * len(fieldlist)

# Create a row for column headings
for col, name in enumerate(fieldlist):
    desc = headings[name]
    ws.write_string(row, col, desc, header_style)
    if name not in colwidths:
        maxwidths[col] = textwidth(desc, bold=True)

breakvalue = None
infile.posf()
while not infile.readn():
    row += 1
    if breakfield:
        if infile[breakfield] != breakvalue and breakvalue is not None:
            row += 1
        breakvalue = infile[breakfield]
    for col, data in enumerate(infile.get(fieldlist)):
        fieldname = fieldlist[col]
        nativedate = False
        nativetime = False
        if infile.fieldType(fieldname) == 'DATE':
            # A native date is read by iSeriesPython as a formatted
            # string.  By default, *ISO format is used, but this can be
            # altered by the DATFMT and DATSEP keywords.  For now,
            # this program only handles *ISO.
            year, month, day = [int(x) for x in data.split('-')]
            if year > 1904:
                ws.write_datetime(
                    row, col, date(year, month, day), date_style)
            nativedate = True
        elif infile.fieldType(fieldname) == 'TIME':
            # A native time is read by iSeriesPython as a formatted
            # string.  By default, *ISO format is used, but this can be
            # altered by the TIMFMT and TIMSEP keywords.  For now,
            # this program only handles *ISO.
            hour, minute, second = [int(x) for x in data.split('.')]
            ws.write_datetime(
                row, col, time(hour, minute, second), time_style)
            nativetime = True
        elif dateflags[fieldname]:
            if data:
                year, md = divmod(data, 10000)
                month, day = divmod(md, 100)
                ws.write_datetime(
                    row, col, date(year, month, day), date_style)
        elif timeflags[fieldname]:
            if data:
                hour, minsec = divmod(data, 10000)
                minute, second = divmod(minsec, 100)
                ws.write_datetime(
                    row, col, time(hour, minute, second), time_style)
        elif data == 0 and fieldname in blankzeros:
            pass
        elif fieldname in numformats:
            ws.write(row, col, data, numformats[fieldname])
        elif fieldname in wrapped:
            ws.write(row, col, data, wrapped_style)
        elif infile.fieldType(fieldname) == 'CHAR':
            ws.write_string(row, col, data, text_style)
        else:
            ws.write(row, col, data)
        if fieldname not in colwidths:
            if nativedate or dateflags[fieldname]:
                maxwidths[col] = datewidth()
            elif nativetime or timeflags[fieldname]:
                maxwidths[col] = timewidth()
            if fieldname in decplaces:
                dp = decplaces[fieldname]
                cf = commaflags[fieldname]
                maxwidths[col] = max(maxwidths[col], numwidth(data, dp, cf))
            else:
                maxwidths[col] = max(maxwidths[col], textwidth(data))
infile.close()

# Set column widths
for col in range(len(fieldlist)):
    if fieldlist[col] in colwidths:
        ws.set_column(col, col, colwidths[fieldlist[col]])
    else:
        ws.set_column(col, col, maxwidths[col])

wb.close()
sndmsg('File copied to ' + sys.argv[2] + '.')
