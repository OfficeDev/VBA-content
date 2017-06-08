---
title: Range Object
keywords: vbagr10.chm5207906
f1_keywords:
- vbagr10.chm5207906
ms.prod: excel
api_name:
- Excel.Range
ms.assetid: 8bc4841b-72f7-34b5-a299-3357bf8f457b
ms.date: 06/08/2017
---


# Range Object

Represents a cell, a row, a column, or a selection of cells that contains one or more contiguous blocks of cells.


## Using the Range Object

The following properties for returning a  **Range** object are described in this section:


-  **Range** property
    
-  **Cells** property
    

## Range Property

Use  **Range**( _arg_), where  _arg_ is the name of the range, to return a **Range** object that represents a single cell or a range of cells. The following example places the value of cell A1 in cell A5.


```
myChart.Application.DataSheet.Range("A5").Value = _ 
    myChart.Application.DataSheet.Range("A1").Value
```

The following example fills the range A1:H8 with the value 20.




```
myChart.Application.DataSheet.Range("A1:H8").Value = 20
```


## Cells Property

Use  **Cells**( _row_,  _column_), where  _row_ is the row's index number and _column_ is the column's index number, to return a single cell. The following example sets the value of cell A1 to 24 (column A is the second column on the datasheet, and row 1 is the second row on the datasheet).


```
myChart.Application.DataSheet.Cells(2, 2).Value = 24
```

Although you can also use  `Range("A1")` to return cell A1, there may be times when the **Cells** property is more convenient because you can use a variable for the row or column. The following example creates column and row headings on the datasheet.




```vb
Sub SetUpTable() 
With myChart.Application.DataSheet 
    For theYear = 1 To 5 
        .Cells(1, theYear + 1).Value = 1990 + theYear 
    Next theYear 
    For theQuarter = 1 To 4 
        .Cells(theQuarter + 1, 1).Value = "Q" &; theQuarter 
    Next theQuarter 
End With 
End Sub
```

Although you can use Visual Basic string functions to alter A1-style references, it's much easier (and much better programming practice) to use the  `Cells(1, 1)` notation.

Use  _expression_. **Cells**( _row_,  _column_), where  _expression_ is an expression that returns a **Range** object, and _row_ and _column_ are relative to the upper-left corner of the range, to return part of a range. The following example sets the value for cell C5.




```
myChart.Application.Range("C5:C10").Cells(1, 1).Value = 35
```


