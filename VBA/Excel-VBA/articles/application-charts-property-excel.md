---
title: Application.Charts Property (Excel)
keywords: vbaxl10.chm132085
f1_keywords:
- vbaxl10.chm132085
ms.prod: excel
api_name:
- Excel.Application.Charts
ms.assetid: d60d912c-7c70-7004-d876-e83b98a13de9
ms.date: 06/08/2017
---


# Application.Charts Property (Excel)

Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the chart sheets in the active workbook.


## Syntax

 _expression_ . **Charts**

 _expression_ A variable that represents an **Application** object.


## Example

This example sets the text for the title of Chart1.


```vb
With Charts("Chart1") 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```

This example hides Chart1, Chart3, and Chart5.




```vb
Charts(Array("Chart1", "Chart3", "Chart5")).Visible = False
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

