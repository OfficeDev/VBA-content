---
title: Application.Selection Property (Excel)
keywords: vbaxl10.chm183107
f1_keywords:
- vbaxl10.chm183107
ms.prod: EXCEL
api_name:
- Excel.Application.Selection
ms.assetid: f25b5608-035b-983a-545d-d720990c28be
---


# Application.Selection Property (Excel)

Returns the selected object in the active window for an  **[Application](application-object-excel.md)** object.


## Syntax

 _expression_ . **Selection**

 _expression_ A variable that represents an **Application** object.


## Remarks

The returned object type depends on the current selection (for example, if a cell is selected, this property returns a  **[Range](range-object-excel.md)** object). The **Selection** property returns **Nothing** if nothing is selected.

Using this property with no object qualifier is equivalent to using  `Application.Selection`.


## Example

This example clears the selection on Sheet1 (assuming that the selection is a range of cells).


```vb
Worksheets("Sheet1").Activate 
Selection.Clear
```

This example displays the Visual Basic object type of the selection.




```vb
Worksheets("Sheet1").Activate 
MsgBox "The selection object type is " &; TypeName(Selection)
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

