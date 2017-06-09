---
title: UpBars Object (Word)
keywords: vbawd10.chm2761
f1_keywords:
- vbawd10.chm2761
ms.prod: word
api_name:
- Word.UpBars
ms.assetid: 22dff1d2-8f1b-8c48-354c-570906e0f830
ms.date: 06/08/2017
---


# UpBars Object (Word)

Represents the up bars in a chart group.


## Remarks

Up bars connect points on series one with higher values on the last series in the chart group (the lines go up from series one). Only 2-D line groups that contain at least two series can have up bars. This object is not a collection. There is no object that represents a single up bar; you either enable up bars for all points in a chart group or you disable them.

If the  **[HasUpDownBars](chartgroup-hasupdownbars-property-word.md)** property is **False** , most properties of the **UpBars** object are disabled.


## Example

Use the  **[UpBars](chartgroup-upbars-property-word.md)** property to return the **UpBars** object. The following example enables up and down bars for chart group one of the first chart in the active document. The example then sets the up bar color to blue and sets the down bar color to red.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .UpBars.Interior.Color = RGB(0, 0, 255) 
 .DownBars.Interior.Color = RGB(255, 0, 0) 
 End With 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


