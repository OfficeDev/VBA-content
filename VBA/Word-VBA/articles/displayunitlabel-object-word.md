---
title: DisplayUnitLabel Object (Word)
keywords: vbawd10.chm1443
f1_keywords:
- vbawd10.chm1443
ms.prod: word
api_name:
- Word.DisplayUnitLabel
ms.assetid: 9b028f6c-fd66-f767-f3d1-09de0fbdc148
ms.date: 06/08/2017
---


# DisplayUnitLabel Object (Word)

Represents a unit label on an axis in the specified chart.


## Remarks

 Unit labels are useful for charting large values (for example, in the millions or billions). You can make the chart more readable by using a single unit label instead of large numbers at each tick mark.


## Example

Use the  **[DisplayUnitLabel](axis-displayunitlabel-property-word.md)** property to return the **DisplayUnitLabel** object. The following example sets the display label caption to "Millions" on the value axis of the first chart in the active document, and then the example turns off automatic font scaling.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .DisplayUnit = xlMillions 
 .HasDisplayUnitLabel = True 
 With .DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
 End With 
 End With 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


