---
title: DataLabels.Separator Property (Word)
keywords: vbawd10.chm207489003
f1_keywords:
- vbawd10.chm207489003
ms.prod: word
api_name:
- Word.DataLabels.Separator
ms.assetid: daf3afde-8a33-de08-a615-57537855818a
ms.date: 06/08/2017
---


# DataLabels.Separator Property (Word)

Sets or returns the separator for the data labels on a chart. Read/write  **Variant** .


## Syntax

 _expression_ . **Separator**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


## Remarks

If you use a string, you will get a string as the separator. If you use  **xlDataLabelSeparatorDefault** (= 1), you will get the default data label separator, which is either a comma or a newline character, depending on the data label.


## Example

The following example sets the data label separator for the first series on the first chart in the active document to a semicolon.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1). _ 
 DataLabels.Separator = ";" 
 End If 
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-word.md)

