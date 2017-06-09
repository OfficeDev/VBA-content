---
title: Axes Object (Word)
ms.prod: word
api_name:
- Word.Axes
ms.assetid: 57261ca9-7fd6-ba99-19bd-5df8e940f714
ms.date: 06/08/2017
---


# Axes Object (Word)

Represents a collection of all the  **[Axis](axis-object-word.md)** objects in the specified chart.


## Remarks

Use the  **[Axes](chart-axes-method-word.md)** method to return the **Axes** collection.

Use  **Axes** ( _Type_ , _AxisGroup_ ), where _Type_ is the axis type and _AxisGroup_ is the axis group, to return an **Axes** collection that contains a single **Axis** object. _Type_ can be one of the following **[XlAxisType](xlaxistype-enumeration-word.md)** constants: **xlCategory** , **xlSeries** , or **xlValue** . _AxisGroup_ can be one of the following **[XlAxisGroup](xlaxisgroup-enumeration-word.md)** constants: **xlPrimary** or **xlSecondary** .


## Example

The following example displays the number of axes for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 MsgBox .Chart.Axes.Count 
 End If 
End With
```

The following example sets the category axis title text for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
 End With 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


