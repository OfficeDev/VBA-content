---
title: DataLabels.Separator Property (Excel)
keywords: vbaxl10.chm584104
f1_keywords:
- vbaxl10.chm584104
ms.prod: excel
api_name:
- Excel.DataLabels.Separator
ms.assetid: bdcd548c-f992-064d-9638-95e1193c3b15
ms.date: 06/08/2017
---


# DataLabels.Separator Property (Excel)

Sets or returns a  **Variant** representing the separator used for the data labels on a chart. Read/write.


## Syntax

 _expression_ . **Separator**

 _expression_ A variable that represents a **DataLabels** object.


## Remarks

If you use a string, you will get a string as the separator. If you use  **xlDataLabelSeparatorDefault** (= 1), you will get the default data label separator, which is either a comma or a newline, depending on the data label.

When a value of "1" is returned, it indicates that the user has not changed the default separator which is a comma ",". You can also pass a value of "1" to change the separator back to the default separator. 

The chart must first be active before you can access the data labels programmatically; otherwise a run-time error occurs.


## Example

This example sets the data label separator for the first series on the first chart to a semicolon. This example assumes a chart exists on the active worksheet.


```vb
Sub ChangeSeparator() 
 
 ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1) _ 
 .DataLabels.Separator = ";" 
 
End Sub
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-excel.md)

