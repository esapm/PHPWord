# Section

Available Section style options:

- ``borderBottomColor``. Border bottom color.
- ``borderBottomSize``. Border bottom size in *twip*.
- ``borderLeftColor``. Border left color.
- ``borderLeftSize``. Border left size in *twip*.
- ``borderRightColor``. Border right color.
- ``borderRightSize``. Border right size in *twip*.
- ``borderTopColor``. Border top color.
- ``borderTopSize``. Border top size in *twip*.
- ``breakType``. Section break type (nextPage, nextColumn, continuous, evenPage, oddPage).
- ``colsNum``. Number of columns.
- ``colsSpace``. Spacing between columns.
- ``footerHeight``. Spacing to bottom of footer.
- ``gutter``. Page gutter spacing.
- ``headerHeight``. Spacing to top of header.
- ``marginTop``. Page margin top in *twip*.
- ``marginLeft``. Page margin left in *twip*.
- ``marginRight``. Page margin right in *twip*.
- ``marginBottom``. Page margin bottom in *twip*.
- ``orientation``. Page orientation (``portrait``, which is default, or ``landscape``).
   See ``\PhpOffice\PhpWord\Style\Section::ORIENTATION_...`` class constants for possible values
- ``pageSizeH``. Page height in *twip*. Implicitly defined by ``orientation`` option. Any changes are discouraged.
- ``pageSizeW``. Page width in *twip*. Implicitly defined by ``orientation`` option. Any changes are discouraged.
- ``vAlign``. Vertical Page Alignment
   See ``\PhpOffice\PhpWord\SimpleType\VerticalJc`` for possible values
- ``pageNumbering``. Array containing page numbering options (start, fmt, chapSep, chapStyle).
   - ``start``. Specifies the page number that appears on the first page of the section. If the value is omitted, numbering continues from the highest page number in the previous section.
   - ``fmt``. Specifies the number format to be used for page numbers in the section. The most commonly used values are:
      - ``cardinalText`` - the cardinal text of the run language. (In English, One, Two, Three, etc.)
      - ``decimal`` - decimal numbering (1, 2, 3, 4, etc.)
      - ``decimalEnclosedCircle`` - decimal number enclosed in a circle
      - ``decimalEnclosedFullstop`` - decimal number followed by a period
      - ``decimalEnclosedParen`` - decimal number enclosed in parentheses
      - ``decimalZero`` - decimal number but with a zero added to numbers 1 through 9
      - ``lowerLetter`` - based on the run language (e.g., a, b, c, etc.). Letters repeat for values greater than the size of the alphabet
      - ``lowerRoman`` - lowercase Roman numerals (i, ii, iii, iv, etc.)
none
      - ``ordinalText`` - ordinal text of the run laguage. (In English, First, Second, Third, etc.)
      - ``upperLetter`` - based on the run language (e.g., A, B, C, etc.). Letters repeat for values greater than the size of the alphabet
      - ``upperRoman`` - uppercase Roman numerals (I, II, III, IV, etc.)
   - ``chapSep``. Specifies the separator character to be used between the chapter and page number if a chapter style has been set. Possible values are:
      - ``colon``  - (e.g., 1:1)
      - ``emDash`` - (e.g., 1—1)
      - ``endash`` - (e.g., 1–1)
      - ``hyphen`` - (e.g., 1-1)
      - ``period`` - (e.g., 1.1)
   - ``chapStyle``. Specifies the one-based index of the heading style applied to chapter titles. The nearest heading of the style is located and the numbering information is extracted for the page number. For example, a chapStyle value of 1 (Heading 1 style) means that the numbering value of the nearest Heading 1 style is used for the chapter value in the page number.
   See ISO/IEC 29500-1 for additional information.