---
title: Refer to Sheets by Name
keywords: vbaxl10.chm5204443
f1_keywords:
- vbaxl10.chm5204443
ms.prod: excel
ms.assetid: 8e58c0d0-ff97-fb00-6afc-f14e2f9c425d
ms.date: 06/08/2017
---


# Refer to Sheets by Name

You can identify sheets by name using the  **[Worksheets](workbook-worksheets-property-excel.md)** and  **[Charts](workbook-charts-property-excel.md)** properties. The following statements activate various sheets in the active workbook.


```vb
Worksheets("Sheet1").Activate 
Charts("Chart1").Activate
```


```vb
DialogSheets("Dialog1").Activate
```

You can use the  **[Sheets](workbook-sheets-property-excel.md)** property to return a worksheet, chart, module, or dialog sheet. The  **Sheets** collection contains all of these kinds of sheets. The following example activates the sheet named "Chart1" in the active workbook.



```vb
Sub ActivateChart() 
 Sheets("Chart1").Activate 
End Sub
```


 **Note**   Charts embedded in a worksheet are members of the **[ChartObjects](chartobjects-object-excel.md)** collection, whereas charts that exist on their own sheets belong to the  **[Charts](charts-object-excel.md)** collection.


