---
title: DataLabel.AutoText Property (Word)
keywords: vbawd10.chm233898119
f1_keywords:
- vbawd10.chm233898119
ms.prod: word
api_name:
- Word.DataLabel.AutoText
ms.assetid: de19c6ef-38a2-0555-49e9-a63b4adb3f72
ms.date: 06/08/2017
---


# DataLabel.AutoText Property (Word)

 **True** if the object automatically generates appropriate text based on context. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoText**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-word.md)** object.


## Example

The following example sets the data labels for series one of the first chart in the active document to automatically generate appropriate text.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1). _ 
 DataLabels.AutoText = True 
 End If 
End With
```


## See also


#### Concepts


[DataLabel Object](datalabel-object-word.md)

