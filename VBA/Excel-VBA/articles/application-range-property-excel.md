---
title: Application.Range Property (Excel)
keywords: vbaxl10.chm183103
f1_keywords:
- vbaxl10.chm183103
ms.prod: excel
api_name:
- Excel.Application.Range
ms.assetid: fec5050e-e6d9-6736-a9bc-b3e7d213a755
ms.date: 06/08/2017
---


# Application.Range Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents a cell or a range of cells.


## Syntax

 _expression_ . **Range**( **_Cell1_** , **_Cell2_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cell1_|Required| **Variant**|The name of the range. This must be an A1-style reference in the language of the macro. It can include the range operator (a colon), the intersection operator (a space), or the union operator (a comma). It can also include dollar signs, but they?re ignored. You can use a local defined name in any part of the range. If you use a name, the name is assumed to be in the language of the macro.|
| _Cell2_|Optional| **Variant**|The cell in the upper-left and lower-right corner of the range. Can be a  **Range** object that contains a single cell, an entire column, or entire row, or it can be a string that names a single cell in the language of the macro.|

## Remarks

When used without an object qualifier, this property is a shortcut for  `ActiveSheet.Range` (it returns a range from the active sheet; if the active sheet isn?t a worksheet, the property fails).

When applied to a  **Range** object, the property is relative to the **Range** object. For example, if the selection is cell C3, then `Selection.Range("B1")` returns cell D3 because it?s relative to the **Range** object returned by the **Selection** property. On the other hand, the code `ActiveSheet.Range("B1")` always returns cell B1.


## Example

This example sets the value of cell A1 on Sheet1 to 3.14159.


```vb
Worksheets("Sheet1").Range("A1").Value = 3.14159
```

This example creates a formula in cell A1 on Sheet1.




```vb
Worksheets("Sheet1").Range("A1").Formula = "=10*RAND()"
```

This example loops on cells A1:D10 on Sheet1. If one of the cells has a value less than 0.001, the code replaces that value with 0 (zero).




```vb
For Each c in Worksheets("Sheet1").Range("A1:D10") 
 If c.Value < .001 Then 
 c.Value = 0 
 End If 
Next c
```

This example loops on the range named "TestRange" and displays the number of empty cells in the range.




```vb
numBlanks = 0 
For Each c In Range("TestRange") 
 If c.Value = "" Then 
 numBlanks = numBlanks + 1 
 End If 
Next c 
MsgBox "There are " &; numBlanks &; " empty cells in this range"
```

This example sets the font style in cells A1:C5 on Sheet1 to italic. The example uses Syntax 2 of the  **Range** property.




```vb
Worksheets("Sheet1").Range(Cells(1, 1), Cells(5, 3)). _ 
 Font.Italic = True 

```


## See also


#### Concepts


[Application Object](application-object-excel.md)

