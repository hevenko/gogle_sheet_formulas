CELL_COLUMN (acell)
=LEFT(ADDRESS(1,COLUMN(acell),4), FIND("1",ADDRESS(1,COLUMN(acell),4)) - 1)
=LEFT(ADDRESS(1;COLUMN(acell);4); FIND("1";ADDRESS(1;COLUMN(acell);4)) - 1)

DYNAMICRANGE (top_left_cell)
=OFFSET(top_left_cell,0,0,1,COLUMN(INDEX(INDIRECT(SHEET_NAME(top_left_cell,true)&ROW(top_left_cell)&":"&ROW(top_left_cell)), MAX(FILTER(COLUMN(INDIRECT(SHEET_NAME(top_left_cell,true)&ROW(top_left_cell)&":"&ROW(top_left_cell))),INDIRECT(SHEET_NAME(top_left_cell,true)&ROW(top_left_cell)&":"&ROW(top_left_cell))  <> "")))) - COLUMN(top_left_cell) + 1)
=OFFSET(top_left_cell;0;0;1;COLUMN(INDEX(INDIRECT(SHEET_NAME(top_left_cell;true)&ROW(top_left_cell)&":"&ROW(top_left_cell)); MAX(FILTER(COLUMN(INDIRECT(SHEET_NAME(top_left_cell;true)&ROW(top_left_cell)&":"&ROW(top_left_cell)));INDIRECT(SHEET_NAME(top_left_cell;true)&ROW(top_left_cell)&":"&ROW(top_left_cell))  <> "")))) - COLUMN(top_left_cell) + 1)

INDIRECT_SHEET_NAME (acell, extended)
=IFERROR(IF(extended,MID(FORMULATEXT(acell), FIND("'", FORMULATEXT(acell)), FIND("!", FORMULATEXT(acell)) - FIND("'", FORMULATEXT(acell)) + 1),MID(FORMULATEXT(acell), FIND("'", FORMULATEXT(acell)) + 1, FIND("!", FORMULATEXT(acell)) - FIND("'", FORMULATEXT(acell)) - 2)),"")
=IFERROR(IF(extended;MID(FORMULATEXT(acell); FIND("'"; FORMULATEXT(acell)); FIND("!"; FORMULATEXT(acell)) - FIND("'"; FORMULATEXT(acell)) + 1);MID(FORMULATEXT(acell); FIND("'"; FORMULATEXT(acell)) + 1; FIND("!"; FORMULATEXT(acell)) - FIND("'"; FORMULATEXT(acell)) - 2));"")

INVOICE_REF (cell)
=MID(FORMULATEXT(cell),2,LEN(FORMULATEXT(cell))-1)
=MID(FORMULATEXT(cell);2;LEN(FORMULATEXT(cell))-1)

RANGECOORDS (raspon)
=CONCATENATE(ADDRESS(ROW(raspon), COLUMN(raspon),4), ":", ADDRESS(ROW(raspon)+ROWS(raspon)-1, COLUMN(raspon)+COLUMNS(raspon)-1, 4))
=CONCATENATE(ADDRESS(ROW(raspon); COLUMN(raspon);4); ":"; ADDRESS(ROW(raspon)+ROWS(raspon)-1; COLUMN(raspon)+COLUMNS(raspon)-1; 4))

SHEET_NAME (acell, extended)
=IFERROR(IF(extended,LEFT(CELL("address", acell), FIND("!", CELL("address", acell))),MID(CELL("address", acell), FIND("'" , CELL("address", acell)) + 1, FIND("!", CELL("address", acell)) - FIND("'" , CELL("address", acell)) - 2)),"")
=IFERROR(IF(extended;LEFT(CELL("address"; acell); FIND("!"; CELL("address"; acell)));MID(CELL("address"; acell); FIND("'" ; CELL("address"; acell)) + 1; FIND("!"; CELL("address"; acell)) - FIND("'" ; CELL("address"; acell)) - 2));"")

WHEN_PAID (INVOICE_REF, COLUMN_HEADER, RANGE_TOP_CELL)
=IFERROR(INDIRECT(SHEET_NAME(RANGE_TOP_CELL,true)&ADDRESS(COLUMN_HEADER,COLUMN(INDIRECT(WHEREIS(INVOICE_REF,DYNAMICRANGE(RANGE_TOP_CELL),SHEET_NAME(RANGE_TOP_CELL,false)))))), "")
=IFERROR(INDIRECT(SHEET_NAME(RANGE_TOP_CELL;true)&ADDRESS(COLUMN_HEADER;COLUMN(INDIRECT(WHEREIS(INVOICE_REF;DYNAMICRANGE(RANGE_TOP_CELL);SHEET_NAME(RANGE_TOP_CELL;false)))))); "")

WHEREIS (CELL, RANGE, SHEET_NAME)
=FINDREFERENCE(ADDRESS(ROW(CELL),COLUMN(CELL),4),RANGECOORDS(RANGE), SHEET_NAME,CELL)
=FINDREFERENCE(ADDRESS(ROW(CELL);COLUMN(CELL);4);RANGECOORDS(RANGE); SHEET_NAME;CELL)

WHICH_INVOICE (ROWHEADER, COLLHEADER, INVOICE)
=CONCATENATE(INDIRECT(CONCATENATE(INDIRECT_SHEET_NAME(INVOICE,true),ROWHEADER,ROW(INDIRECT(INVOICE_REF(INVOICE))))), " ",INDIRECT(ADDRESS(COLLHEADER,COLUMN(INDIRECT(INVOICE_REF(INVOICE))),1,true,INDIRECT_SHEET_NAME(INVOICE,false))), " mj.")
=CONCATENATE(INDIRECT(CONCATENATE(INDIRECT_SHEET_NAME(INVOICE;true);ROWHEADER;ROW(INDIRECT(INVOICE_REF(INVOICE))))); " ";INDIRECT(ADDRESS(COLLHEADER;COLUMN(INDIRECT(INVOICE_REF(INVOICE)));1;true;INDIRECT_SHEET_NAME(INVOICE;false))); " mj.")