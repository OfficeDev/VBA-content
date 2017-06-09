---
title: ChartData.IsLinked Property (Word)
keywords: vbawd10.chm190382082
f1_keywords:
- vbawd10.chm190382082
ms.prod: word
api_name:
- Word.ChartData.IsLinked
ms.assetid: d22ba8ec-2e6e-aa46-6e4f-a370a01d0835
ms.date: 06/08/2017
---


# ChartData.IsLinked Property (Word)

 **True** if the data for the chart is linked to an external Microsoft Excel workbook. Read-only **Boolean** .


## Syntax

 _expression_ . **IsLinked**

 _expression_ A variable that represents a **[ChartData](chartdata-object-word.md)** object.


## Remarks

Using the  **[BreakLink](chartdata-breaklink-method-word.md)** method to remove the link to an Excel workbook sets this property to **False** .


## Example

The following example verifies whether the data for the first chart in the active document is linked to an external Excel workbook. If the data for the chart is linked, the example then uses the  **BreakLink** method to remove the link. If the data for the chart is not linked, the example uses the **[Activate](chartdata-activate-method-word.md)** method to display the embedded data for the chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartData 
 If .IsLinked Then 
 .BreakLink 
 Else 
 .Activate 
 End If 
 End With 
 End If 
End With
```


## See also


#### Concepts


[ChartData Object](chartdata-object-word.md)

