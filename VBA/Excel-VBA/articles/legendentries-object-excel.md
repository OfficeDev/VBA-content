---
title: LegendEntries Object (Excel)
keywords: vbaxl10.chm587072
f1_keywords:
- vbaxl10.chm587072
ms.prod: excel
api_name:
- Excel.LegendEntries
ms.assetid: 51d98149-b90b-432b-7771-0815a0e89655
ms.date: 06/08/2017
---


# LegendEntries Object (Excel)

A collection of all the  **[LegendEntry](legendentry-object-excel.md)** objects in the specified chart legend.


## Remarks

 Each legend entry has two parts: the text of the entry, which is the name of the series or trendline associated with the legend entry; and the entry marker, which visually links the legend entry with its associated series or trendline in the chart. The formatting properties for the entry marker and its associated series or trendline are contained in the **[LegendKey](legendkey-object-excel.md)** object.


## Example

Use the  **[LegendEntries](legend-legendentries-method-excel.md)** method to return the **LegendEntries** collection. The following example loops through the collection of legend entries in embedded chart one and changes their font color.


```
With Worksheets("sheet1").ChartObjects(1).Chart.Legend 
 For i = 1 To .LegendEntries.Count 
 .LegendEntries(i).Font.ColorIndex = 5 
 Next 
End With
```

Use  **[LegendEntries](legend-legendentries-method-excel.md)** ( _index_ ), where _index_ is the legend entry index number, to return a single **LegendEntry** object. You cannot return legend entries by name.



The index number represents the position of the legend entry in the legend.  `LegendEntries(1)` is at the top of the legend; `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font style for the text of the legend entry at the top of the legend (this is usually the legend for series one) in embedded chart one to italic.




```
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Legend.LegendEntries(1).Font.Italic = True
```


## Methods



|**Name**|
|:-----|
|[Item](legendentries-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](legendentries-application-property-excel.md)|
|[Count](legendentries-count-property-excel.md)|
|[Creator](legendentries-creator-property-excel.md)|
|[Parent](legendentries-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
