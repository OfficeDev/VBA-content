---
title: Chart.CategoryLabelLevel Property (Excel)
keywords: vbaxl10.chm149195
f1_keywords:
- vbaxl10.chm149195
ms.prod: excel
ms.assetid: b3a54685-18d7-8c24-b2e8-f3bfb03fc69e
ms.date: 06/08/2017
---


# Chart.CategoryLabelLevel Property (Excel)

Returns a  **[XlCategoryLabelLevel Enumeration (Excel)](xlcategorylabellevel-enumeration-excel.md)** constant referring to the level of where the category labels are being sourced from. **Integer** Read/Write.


## Syntax

 _expression_ . **CategoryLabelLevel**

 _expression_ A variable that represents a **Chart** object.


## Remarks

If there is a hierarchy, 0 refers to the most parent level, 1 refers to its children, and so on. So, 0 equals the first level, 1 equals the second level, 2 equals the third level, and so on. 


## Example

The following sample code uses the  **CategoryNameLevel** property to set the chart categories names from previously created row.


```vb
    Sheets(1).Range("C1:E1").Value2 = "Sample_Row1"
    Sheets(1).Range("C2:E2").Value2 = "Sample_Row2"
    Sheets(1).Range("A3:A5").Value2 = "Sample_ColA"
    Sheets(1).Range("B3:B5").Value2 = "Sample_ColB"
    Sheets(1).Range("C3:E5").Formula = "=row()"
    Dim crt As Chart
    Set crt = Sheets(1).ChartObjects.Add(0, 0, 500, 200).Chart
    crt.SetSourceData Sheets(1).Range("A1:E5")
    ' Set the series names to only use column B
    crt.SeriesNameLevel = 1
    ' Use columns A and B for the series names
    crt.SeriesNameLevel = xlSeriesNameLevelAll
    ' Use row 1 for the category labels
    crt.CategoryLabelLevel = 0
    ' Use rows 1 and 2 for the category labels
    crt.CategoryLabelLevel = xlCategoryLabelLevelAll
```


## Property value

 **XLCATEGORYLABELLEVEL**


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

