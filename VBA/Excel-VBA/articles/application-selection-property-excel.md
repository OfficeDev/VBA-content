---
title: Application.Selection Property (Excel)
keywords: vbaxl10.chm183107
f1_keywords:
- vbaxl10.chm183107
ms.prod: excel
api_name:
- Excel.Application.Selection
ms.assetid: f25b5608-035b-983a-545d-d720990c28be
ms.date: 04/25/2018
---


# Application.Selection Property (Excel)

Returns the currently selected object in the active worksheet for an **[Application](application-object-excel.md)** object. Returns `Nothing` if no objects are selected. Use the `Select` method to set the selection, and use `TypeName` to discover the kind of object that is selected. 


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
This example displays information about the current selection.

```vb
Sub TestSelection(  )
    Dim str As String
    Select Case TypeName(Selection)
    Case "Nothing"
        str = "No selection made."
    Case "Range"
        str = "You selected the range: " & Selection.Address
    Case "Picture"
        str = "You selected a picture."
    Case Else
        str = "You selected a " & TypeName(Selection) & "."
    End Select
    MsgBox str
End Sub
```

## See also

[TypeName function](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/typename-function)<br>
[Application Object](application-object-excel.md)

