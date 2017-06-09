---
title: Chart.PlotBy Property (PowerPoint)
keywords: vbapp10.chm65738
f1_keywords:
- vbapp10.chm65738
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.PlotBy
ms.assetid: 14b696d7-148c-267f-4294-4dddc9fba4e1
ms.date: 06/08/2017
---


# Chart.PlotBy Property (PowerPoint)

Returns or sets the way columns or rows are used as data series on the chart. Read/write  **Long**.


## Syntax

 _expression_. **PlotBy**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Remarks

The value of this property can be one of the following  **[XlRowCol](xlrowcol-enumeration-powerpoint.md)** constants:


-  **xlColumns**
    
-  **xlRows**
    


For PivotChart reports, this property is read-only and always returns  **xlColumns**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes the first chart in the active document to plot data by columns.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.PlotBy = xlColumns

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

