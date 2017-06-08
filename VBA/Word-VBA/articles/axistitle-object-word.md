---
title: AxisTitle Object (Word)
keywords: vbawd10.chm1499
f1_keywords:
- vbawd10.chm1499
ms.prod: word
api_name:
- Word.AxisTitle
ms.assetid: ec746a05-40df-95cc-c017-40ef150504cf
ms.date: 06/08/2017
---


# AxisTitle Object (Word)

Represents a chart axis title.


## Remarks

Use the  **[AxisTitle](axis-axistitle-property-word.md)** property to return an **AxisTitle** object.

The  **AxisTitle** object does not exist and cannot be used unless the **[HasTitle](axis-hastitle-property-word.md)** property for the axis is **True** .


## Example

The following example sets the caption, sets the font to Bookman 10 point, and formats the word "millions" as italic for the axis title of the value axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .HasTitle = True 
 With .AxisTitle 
 .Caption = "Revenue (millions)" 
 .Font.Name = "bookman" 
 .Font.Size = 10 
 .Characters(10, 8).Font.Italic = True 
 End With 
 End With 
 End If 
End With 

```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


