'''Copy data from a physical file to an Excel binary file in the IFS.

Written by John Yeung.  Last modified 2015-04-29.

Usage (from CL):
    python233/python '/util/cpytoxlsf.py' parm(&pf &xls [&A1text &A2text ...])

The above assumes this program is located in '/util', and that iSeriesPython
2.3.3 is installed.  If you are at V5R3 or later, you should be using
iSeriesPython 2.7 instead.

Some features/caveats:

-  Column headings come from the COLHDG values in the DDS.  Multiple
    values for a single field are joined by spaces, not newlines.  For
    any fields without a COLHDG, or with only blanks in the COLHDG (these
    two situations are indistinguishable), the field name is used as the
    heading (the TEXT keyword is not checked).  To specify a blank column
    heading rather than the field name, use COLHDG('*BLANK').
-  Column headings wrap and are displayed in bold.
-  Each column is sized approximately according to its longest data,
    assuming that the default font is Arial 10, unless a width is
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
'''

import sys
import re
from os import system
from datetime import date, time

# Third-party package available at <http://pypi.python.org/pypi/xlwt>
import xlwt
# If running iSeriesPython 2.3.3, xlwt needs to be modified to work with
# EBCDIC.  This is not required for iSeriesPython 2.5 or 2.7.

# Table of empirically determined character widths in Arial 10, the default
# font used by xlwt.  Note that font rendering is somewhat dependent on the
# configuration of the PC that is used to open the resulting file, so these
# widths are not necessarily exact for other people's PCs.  Also note that
# only characters which are different in width than '0' are needed here.
charwidths = {
    '0': 262.637,
    'f': 146.015,
    'i': 117.096,
    'j': 88.178,
    'k': 233.244,
    'l': 88.178,
    'm': 379.259,
    'r': 175.407,
    's': 233.244,
    't': 117.096,
    'v': 203.852,
    'w': 321.422,
    'x': 203.852,
    'z': 233.244,
    'A': 321.422,
    'B': 321.422,
    'C': 350.341,
    'D': 350.341,
    'E': 321.422,
    'F': 291.556,
    'G': 350.341,
    'H': 321.422,
    'I': 146.015,
    'K': 321.422,
    'M': 379.259,
    'N': 321.422,
    'O': 350.341,
    'P': 321.422,
    'Q': 350.341,
    'R': 321.422,
    'S': 321.422,
    'U': 321.422,
    'V': 321.422,
    'W': 496.356,
    'X': 321.422,
    'Y': 321.422,
    ' ': 146.015,
    '!': 146.015,
    '"': 175.407,
    '%': 438.044,
    '&': 321.422,
    '\'': 88.178,
    '(': 175.407,
    ')': 175.407,
    '*': 203.852,
    '+': 291.556,
    ',': 146.015,
    '-': 175.407,
    '.': 146.015,
    '/': 146.015,
    ':': 146.015,
    ';': 146.015,
    '<': 291.556,
    '=': 291.556,
    '>': 291.556,
    '@': 496.356,
    '[': 146.015,
    '\\': 146.015,
    ']': 146.015,
    '^': 203.852,
    '`': 175.407,
    '{': 175.407,
    '|': 146.015,
    '}': 175.407,
    '~': 291.556}

ezxf = xlwt.easyxf

# I have a custom SNDMSG wrapper that I use to receive immediate messages
# from iSeriesPython; but for basic use, simply printing the message works.
# iSeriesPython also comes with os400.sndmsg, which is essentially a wrapper
# for QMHSNDM.
def sndmsg(msg):
    print msg

def _integer_digits(n):
    '''Return the number of digits in a positive integer'''
    if n == 0:
        return 1
    digits = 0
    while n:
        digits += 1
        n //= 10
    return digits

def number_analysis(n, dp=0):
    '''Return a 4-tuple of (digits, thousands, points, signs)'''
    digits, thousands, points, signs = 0, 0, 0, 0
    if n < 0:
        signs = 1
        n = -n
    if dp > 0:
        points = 1
    if isinstance(n, float):
        idigits = _integer_digits(int(n) + 1)
    elif isinstance(n, (int, long)):
        idigits = _integer_digits(n)
    else:
        return None
    digits = idigits + dp
    thousands = (idigits - 1) // 3
    return digits, thousands, points, signs

def colwidth(n):
    '''Translate human-readable units to BIFF column width units'''
    if n <= 0:
        return 0
    if n <= 1:
        return int(n * 456)
    return int(200 + n * 256)

def fitwidth(data, bold=False):
    '''Try to autofit Arial 10'''
    units = 220
    for char in str(data):
        if char in charwidths:
            units += charwidths[char]
        else:
            units += charwidths['0']
    if bold:
        units *= 1.1
    return int(max(units, 700))  # Don't go smaller than a reported width of 2

def numwidth(data, dp, use_commas=False):
    '''Try to autofit a number in Arial 10'''
    units = 220
    digits, commas, points, signs = number_analysis(data, dp)
    units += digits * charwidths['0']
    if use_commas:
        units += commas * charwidths[',']
    units += points * charwidths['.']
    units += signs * charwidths['-']
    return int(max(units, 700))  # Don't go smaller than a reported width of 2

def datewidth():
    return int(220 + 8 * charwidths['0'] + 2 * charwidths['/'])

def timewidth():
    digits_width = 6 * charwidths['0']
    sep_width = 2 * charwidths[':']
    space_width = charwidths[' ']
    am_pm_width = max(charwidths['A'], charwidths['P']) + charwidths['M']
    return int(220 + digits_width + sep_width + space_width + am_pm_width)

def default_numformat(dp=0, use_commas=False):
    '''Generate a style object for Excel fixed number format'''
    integers, decimals = '0', ''
    if use_commas:
        integers = '#,##0'
    if dp > 0:
        decimals = '.' + '0' * dp
    combined = integers + decimals
    return ezxf(num_format_str=combined)

def editcode(code, dp=0):
    '''Generate a style object corresponding to an edit code'''
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
    return ezxf(num_format_str=';'.join((positive, negative, zero)))

def is_numeric_date(size, editword):
    return size == (8, 0) and editword in ("'    -  -  '", "'    /  /  '")

def is_numeric_time(size, editword):
    return size == (6, 0) and editword in ("'  .  .  '", "'  :  :  '")

# Check parameters
parameters = len(sys.argv) - 1
if parameters < 2:
    sndmsg('Program needs at least 2 parameters; received %d.' % parameters)
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

infile = File400(filename, 'r', lib=libname)
if libname.startswith('*'):
    libname = infile.libName()
sndmsg('Opened ' + libname + '/' + filename + ' for reading.')

# Get column headings and formatting information from the DDS
fieldlist = []
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
template = "dspffd %s/%s output(*outfile) outfile(qtemp/dspffdpf)"
system(template % (libname, filename))
ddsfile = File400('DSPFFDPF', 'r', lib='QTEMP')

ddsfile.posf()
while not ddsfile.readn():
    fieldname = ddsfile['WHFLDE']
    fieldtext = ddsfile['WHFTXT']
    rcdtext = ddsfile['WHTEXT']
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
        numformat = ezxf(num_format_str=match.group(1))
    elif numeric:
        numformat = editcode(ddsfile['WHECDE'], ddsfile['WHFLDP'])
    else:
        numformat = None
    if numformat:
        numformats[fieldname] = numformat
        commaflags[fieldname] = ',' in numformat.num_format_str

    # Check whether it looks like a numeric date or time
    dateflags[fieldname] = is_numeric_date(fieldsize, ddsfile['WHEWRD'])
    timeflags[fieldname] = is_numeric_time(fieldsize, ddsfile['WHEWRD'])

    # Look for fixed column width
    match = re.search(r'width=([1-9][0-9]*)', fieldtext, re.IGNORECASE)
    if match:
        colwidths[fieldname] = colwidth(int(match.group(1)))

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
wb = xlwt.Workbook()
ws = wb.add_sheet(infile.fileName())
row = 0

title_style = ezxf('font: bold on')
header_style = ezxf('font: bold on; align: wrap on')
date_style = ezxf(num_format_str='m/d/yyyy')
time_style = ezxf(num_format_str='h:mm:ss AM/PM')
text_style = ezxf(num_format_str='@')
wrapped_style = ezxf('align: wrap on')

# Populate first few rows using additional parameters, if provided.
# Typically, these rows would be used for report ID, date, and title.
for arg in sys.argv[3:]:
    ws.write(row, 0, arg, title_style)
    row += 1

# Keep track of the widest data in each column
maxwidths = [0] * len(fieldlist)

# If there were top-row parameters, skip a row before starting the
# column headings.
if parameters > 2:
    row += 1

# Create a row for column headings
for col, name in enumerate(fieldlist):
    desc = headings[name]
    ws.write(row, col, desc, header_style)
    if name not in colwidths:
        maxwidths[col] = fitwidth(desc, bold=True)

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
            year, month, day = [int(x) for x in data.split('-')]
            if year > 1904:
                ws.write(row, col, date(year, month, day), date_style)
            nativedate = True
        elif infile.fieldType(fieldname) == 'TIME':
            hour, minute, second = [int(x) for x in data.split('.')]
            ws.write(row, col, time(hour, minute, second), time_style)
            nativetime = True
        elif dateflags[fieldname]:
            if data:
                year, md = divmod(data, 10000)
                month, day = divmod(md, 100)
                ws.write(row, col, date(year, month, day), date_style)
        elif timeflags[fieldname]:
            if data:
                hour, minsec = divmod(data, 10000)
                minute, second = divmod(minsec, 100)
                ws.write(row, col, time(hour, minute, second), time_style)
        elif data == 0 and fieldname in blankzeros:
            pass
        elif fieldname in numformats:
            ws.write(row, col, data, numformats[fieldname])
        elif fieldname in wrapped:
            ws.write(row, col, data, wrapped_style)
        elif infile.fieldType(fieldname) == 'CHAR':
            ws.write(row, col, data, text_style)
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
                maxwidths[col] = max(maxwidths[col], fitwidth(data))
infile.close()

# Set column widths
for col in range(len(fieldlist)):
    if fieldlist[col] in colwidths:
        ws.col(col).width = colwidths[fieldlist[col]]
    else:
        ws.col(col).width = maxwidths[col]

wb.save(sys.argv[2])
sndmsg('File copied to ' + sys.argv[2] + '.')
