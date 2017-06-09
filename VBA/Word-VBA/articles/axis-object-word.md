---
title: Axis Object (Word)
keywords: vbawd10.chm1725
f1_keywords:
- vbawd10.chm1725
ms.prod: word
api_name:
- Word.Axis
ms.assetid: 3a7ad7d8-d397-a79a-eb6a-a5f0822cbe5d
ms.date: 06/08/2017
---


# Axis Object (Word)

Represents a single axis in a chart.


## Remarks

The  **Axis** object is a member of the **[Axes](axes-object-word.md)** collection.

Use  **Axes** ( _Type_ , _AxisGroup_ ) where _Type_ is the axis type and _AxisGroup_ is the axis group to return a single **Axis** object. _Type_ can be one of the following **[XlAxisType](xlaxistype-enumeration-word.md)** constants: **xlCategory** , **xlSeries** , or **xlValue** . _AxisGroup_ can be one of the following **[XlAxisGroup](xlaxisgroup-enumeration-word.md)** constants: **xlPrimary** or **xlSecondary** . For more information, see the **[Axes](chart-axes-method-word.md)** method.


## Example

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

