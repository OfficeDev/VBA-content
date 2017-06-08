---
title: ErrorBars Object (Word)
keywords: vbawd10.chm1142
f1_keywords:
- vbawd10.chm1142
ms.prod: word
api_name:
- Word.ErrorBars
ms.assetid: 33949dd1-48fd-9fff-0bec-1439b65d8e04
ms.date: 06/08/2017
---


# ErrorBars Object (Word)

Represents the error bars on a chart series.


## Remarks

 Error bars indicate the degree of uncertainty for chart data. Only series in area, bar, column, line, and scatter groups on a 2-D chart can have error bars. Only series in scatter groups can have x and y error bars. This object is not a collection. There is no object that represents a single error bar; you either enable x error bars or y error bars for all points in a series or you disable them.

The  **[ErrorBar](series-errorbar-method-word.md)** method changes the error bar format and type.


## Example

Use the  **[ErrorBars](series-errorbars-property-word.md)** property to return the **ErrorBars** object. The following example enables error bars for series one of the first chart in the active document and then sets the end style for the error bars.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).HasErrorBars = True 
 .Chart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


