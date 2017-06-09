---
title: Chart.PlotBy Property (Word)
keywords: vbawd10.chm79364298
f1_keywords:
- vbawd10.chm79364298
ms.prod: word
api_name:
- Word.Chart.PlotBy
ms.assetid: ae2774d0-0f58-2224-9104-61d00fa63a86
ms.date: 06/08/2017
---


# Chart.PlotBy Property (Word)

Returns or sets the way columns or rows are used as data series on the chart. Read/write  **Long** .


## Syntax

 _expression_ . **PlotBy**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Remarks

The value of this property can be one of the following  **[XlRowCol](xlrowcol-enumeration-word.md)** constants:


-  **xlColumns**
    
-  **xlRows**
    


For PivotChart reports, this property is read-only and always returns  **xlColumns** .


## Example

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


[Chart Object](chart-object-word.md)

