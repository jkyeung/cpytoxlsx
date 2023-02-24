# README #

`cpytoxlsf.py` and `cpytoxlsx.py` are modules for [iSeriesPython](http://www.iseriespython.com).  The former generates .xls files from physical or logical files and requires [xlwt](https://pypi.python.org/pypi/xlwt); the latter generates .xlsx and requires [XlsxWriter](https://pypi.python.org/pypi/XlsxWriter).

### Module docstring (from cpytoxlsx.py) ###
Copy data from a physical or logical file to an Excel binary file.

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
    by specifying 'break on \<fieldname\>' in the record (not field!) text.

The motivation for this program is to provide a tool for easy generation
of formatted spreadsheets.

Nice-to-have features not yet implemented include general edit word
support (not just for date detection), more available edit codes, more
comprehensive date support, the ability to choose fonts, and automatic
population of multiple sheets given multiple file members.

Also, it would be nice to wrap this in a command for even greater ease of
use, including meaningful promptability.
