---
title: ControlFormat.ListFillRange Property (Excel)
keywords: vbaxl10.chm630082
f1_keywords:
- vbaxl10.chm630082
ms.prod: excel
api_name:
- Excel.ControlFormat.ListFillRange
ms.assetid: 1004b4a7-9315-7736-a71b-1d94d229fd7e
ms.date: 06/08/2017
---


# ControlFormat.ListFillRange Property (Excel)

Returns or sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box. Read/write  **String** .


## Syntax

 _expression_ . **ListFillRange**

 _expression_ A variable that represents a **ControlFormat** object.


## Remarks

Microsoft Excel reads the contents of every cell in the range and inserts the cell values into the list box. The list tracks changes in the range's cells.

If the list in the list box was created with the  **[AddItem](controlformat-additem-method-excel.md)** method, this property returns an empty string ("").


## Example

This example adds a list box to worksheet one and sets the fill range for the list box.


```vb
With Worksheets(1) 
 Set lb = .Shapes.AddFormControl(xlListBox, 100, 10, 100, 100) 
 lb.ControlFormat.ListFillRange = "A1:A10" 
End With
```


## See also


#### Concepts


[ControlFormat Object](controlformat-object-excel.md)

