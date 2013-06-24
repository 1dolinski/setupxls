SetupXls
=============

SetupXls is a small class that allows you to quickly and easily link the power of ruby to Excel.

Usage
-----

Install using `gem install setupxls`

    x = SetupXls.setsheet("[an open Excel workbook sheetname]")

    x.cells(1,2).value = 5 #prints 5 in the worksheets range of B1. 
    y = x.cells(1,2).value #sets the variable y to the value in B1. 

    Remember it's alawys (rw, col)