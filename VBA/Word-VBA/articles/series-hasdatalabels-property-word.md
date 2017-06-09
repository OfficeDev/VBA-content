---
title: Series.HasDataLabels Property (Word)
keywords: vbawd10.chm123732046
f1_keywords:
- vbawd10.chm123732046
ms.prod: word
api_name:
- Word.Series.HasDataLabels
ms.assetid: 2e5ffc2d-11ae-2ab3-a642-fc0349ff356b
ms.date: 06/08/2017
---


# Series.HasDataLabels Property (Word)

 **True** if the series has data labels. Read/write **Boolean** .


## Syntax

 _expression_ . **HasDataLabels**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Example

The following example enables data labels for series three of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(3) 
 .HasDataLabels = True 
 .ApplyDataLabels Type:=xlValue 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

