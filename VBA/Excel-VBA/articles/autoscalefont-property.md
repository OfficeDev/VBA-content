---
title: AutoScaleFont Property
keywords: vbagr10.chm5207069
f1_keywords:
- vbagr10.chm5207069
ms.prod: excel
api_name:
- Excel.AutoScaleFont
ms.assetid: cb21d2e7-d3b9-e135-03ba-6d45275d4590
ms.date: 06/08/2017
---


# AutoScaleFont Property

 **True** if the text in the object changes font size when the object size changes. The default value is **True**. Read/write  **Variant**.


## Example

This example adds a title to the chart, and it causes the title font to remain the same size whenever the chart size changes.


```vb
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "1996 sales" 
 .ChartTitle.AutoScaleFont = False 
End With 

```


