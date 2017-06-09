---
title: Refer to Cells and Ranges by Using A1 Notation
keywords: vbaxl10.chm5204421
f1_keywords:
- vbaxl10.chm5204421
ms.prod: excel
ms.assetid: c98741c5-465e-137f-872d-185a20068d4a
ms.date: 06/08/2017
---


# Refer to Cells and Ranges by Using A1 Notation

You can refer to a cell or range of cells in the A1 reference style by using the  **Range** property. The following subroutine changes the format of cells A1:D5 to bold.


```vb
Sub FormatRange() 
 Workbooks("Book1").Sheets("Sheet1").Range("A1:D5") _ 
 .Font.Bold = True 
End Sub
```


The following table illustrates some A1-style references using the  **Range** property.



|**Reference**|**Meaning**|
|:-----|:-----|
| `Range("A1")`|Cell A1|
| `Range("A1:B5")`|Cells A1 through B5|
| `Range("C5:D9,G9:H16")`|A multiple-area selection|
| `Range("A:A")`|Column A|
| `Range("1:1")`|Row 1|
| `Range("A:C")`|Columns A through C|
| `Range("1:5")`|Rows 1 through 5|
| `Range("1:1,3:3,8:8")`|Rows 1, 3, and 8|
| `Range("A:A,C:C,F:F")`|Columns A, C, and F|

