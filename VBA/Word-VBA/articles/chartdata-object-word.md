---
title: ChartData Object (Word)
keywords: vbawd10.chm2905
f1_keywords:
- vbawd10.chm2905
ms.prod: word
api_name:
- Word.ChartData
ms.assetid: 323ee62c-9b70-8280-d448-79cf4d2b6953
ms.date: 06/08/2017
---


# ChartData Object (Word)

Represents access to the linked or embedded data associated with a chart.


## Remarks

Use the  **[ChartData](chart-chartdata-property-word.md)** property to return the **ChartData** object.


## Example

The following example uses the  **[Activate](chartdata-activate-method-word.md)** method to display the data associated with the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1).Chart.ChartData 
 .Activate 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


