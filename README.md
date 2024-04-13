# README #

`cpytoxlsf.py` and `cpytoxlsx.py` are modules for [iSeriesPython](http://www.iseriespython.com).  The former generates .xls files from physical or logical files and requires [xlwt](https://pypi.python.org/pypi/xlwt); the latter generates .xlsx and requires [XlsxWriter](https://pypi.python.org/pypi/XlsxWriter).

`cpytoxlsx3.py` is a port of `cpytoxlsx.py` to IBM's Python for PASE.  It requires [pyodbc](https://pypi.python.org/pypi/pyodbc) as well as XlsxWriter.

### Module docstring (from cpytoxlsx3.py) ###
Copy data from a physical or logical file to an Excel binary file.

Written by John Yeung.  Last modified 2024-04-12.

Usage (assuming Richard Schoen's QshOni is installed):<br>
&emsp;qshoni/qshpyrun &script_lib 'cpytoxlsx3.py' (<br>
&emsp;&emsp;&pf<br>
&emsp;&emsp;&xlsx<br>
&emsp;&emsp;[&A1_text<br>
&emsp;&emsp;&A2_text<br>
&emsp;&emsp;...])<br>
&emsp;&emsp;&py_version

Required parameters:
  - &script_lib = IFS directory containing cpytoxlsx3.py
  - &pf = qualified name of file to copy
  - &xlsx = name of workbook to create, including path and extension
  - &py_version = Python version; must be at least 3.6

Optional parameters:
  - &A1_text = free-form text to appear at the top of the spreadsheet
  - &A2_text = free-form text to appear on the 2nd line of the spreadsheet

_If using QshOni, the limit is 38 of these optional parameters (because
QSHPYRUN provides up to 40 total arguments to the Python script), with
each limited to 200 characters._

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
