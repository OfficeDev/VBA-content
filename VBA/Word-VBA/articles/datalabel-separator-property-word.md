---
title: DataLabel.Separator Property (Word)
keywords: vbawd10.chm233900011
f1_keywords:
- vbawd10.chm233900011
ms.prod: word
api_name:
- Word.DataLabel.Separator
ms.assetid: 4f681807-d9ec-8c12-585b-6f7bbcb105be
ms.date: 06/08/2017
---


# DataLabel.Separator Property (Word)

Returns or sets the separator used for the data labels on a chart. Read/write  **Variant** .


## Syntax

 _expression_ . **Separator**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-word.md)** object.


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


[DataLabel Object](datalabel-object-word.md)

