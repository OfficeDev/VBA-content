---
title: Worksheet.Range Property (Excel)
keywords: vbaxl10.chm175120
f1_keywords:
- vbaxl10.chm175120
ms.prod: excel
api_name:
- Excel.Worksheet.Range
ms.assetid: 9a323305-c822-ef9e-1cc8-ec077a976834
ms.date: 06/08/2017
---


# Worksheet.Range Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents a cell or a range of cells.


## Syntax

 _expression_ . **Range** ( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|A **String** that is a range reference when one argument is used. Either a **String** that is a range reference or a **Range** object when two arguments are used.|
| _Arg2_|Optional| **Variant**|Either a **String** that is a range reference or a **Range** object. _Arg2_ defines another extremity of the range returned by the property.|


## Remarks

_Arg1_ and _Arg2_ can be A1-style references in the language of the macro. The range references can include the range operator (a colon), intersection operator (a space), or union operator (a comma). They can also include dollar signs, which are ignored. A local defined name can be a range reference. If you use a name, the name is assumed to be in the language of the macro.   

_Arg1_ and _Arg2_ can be **Range** objects that contain a single cell, an entire column, or entire row, or any other range of cells.

Often, _Arg1_ and _Arg2_ are single cells in the upper-left and lower-right corner of the range returned.

When used without an object qualifier, this property is a shortcut for  `ActiveSheet.Range` (it returns a range from the active sheet; if the active sheet isn?t a worksheet, the property fails).

When applied to a  **Range** object, the property is relative to the **Range** object. For example, if the selection is cell C3, then `Selection.Range("B1")` returns cell D3 because it is relative to the **Range** object returned by the **Selection** property. On the other hand, the code `ActiveSheet.Range("B1")` always returns cell B1.


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

This example compares the **Worksheet.Range** property, **[Application.Union](application-union-method-excel.md)** method, and **[Application.Intersect](application-intersect-method-excel.md)** method.

```vb
Range("A1:A10").Select                            'Selects cells A1 to A10.
Range(Range("A1"), Range("A10")).Select           'Selects cells A1 to A10.

Range("A1, A10").Select                           'Selects cells A1 and A10.
Union(Range("A1"), Range("A10")).Select           'Selects cells A1 and A10.

Range("A1:A5 A5:A10").Select                      'Selects cell A5.
Intersect(Range("A1:A5"), Range("A5:A10")).Select 'Selects cell A5.
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

