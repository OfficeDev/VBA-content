---
title: Application.ActiveSheet Property (Excel)
keywords: vbaxl10.chm132078
f1_keywords:
- vbaxl10.chm132078
ms.prod: excel
api_name:
- Excel.Application.ActiveSheet
ms.assetid: 6ed42d87-2ad5-eecc-ad5b-4c92617a04bc
ms.date: 06/08/2017
---


# Application.ActiveSheet Property (Excel)

Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns  **Nothing** if no sheet is active.


## Syntax

 _expression_ . **ActiveSheet**

 _expression_ A variable that represents an **Application** object.


## Remarks

If you do not specify an object qualifier, this property returns the active sheet in the active workbook.

If a workbook appears in more than one window, the  **ActiveSheet** property may be different in different windows.


## Example

This example displays the name of the active sheet.


```vb
MsgBox "The name of the active sheet is " &; ActiveSheet.Name
```

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1)

This example creates a print preview of the active sheet that has the page number at the top of column B on each page.




```vb
Sub PrintSheets()

   'Set up your variables.
   Dim iRow As Integer, iRowL As Integer, iPage As Integer
   'Find the last row that contains data.
   iRowL = Cells(Rows.Count, 1).End(xlUp).Row
   
   'Define the print area as the range containing all the data in the first two columns of the current worksheet.
   ActiveSheet.PageSetup.PrintArea = Range("A1:B" &; iRowL).Address
   
   'Select all the rows containing data.
   Rows(iRowL).Select
   
   'display the automatic page breaks
   ActiveSheet.DisplayAutomaticPageBreaks = True
   Range("B1").Value = "Page 1"
   
   'After each page break, go to the next cell in column B and write out the page number.
   For iPage = 1 To ActiveSheet.HPageBreaks.Count
      ActiveSheet.HPageBreaks(iPage) _
         .Location.Offset(0, 1).Value = "Page " &; iPage + 1
   Next iPage
   
   'Show the print preview, and afterwards remove the page numbers from column B.
   ActiveSheet.PrintPreview
   Columns("B").ClearContents
   Range("A1").Select
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Application Object](application-object-excel.md)

