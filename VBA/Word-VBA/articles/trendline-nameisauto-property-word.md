---
title: Trendline.NameIsAuto Property (Word)
keywords: vbawd10.chm26345660
f1_keywords:
- vbawd10.chm26345660
ms.prod: word
api_name:
- Word.Trendline.NameIsAuto
ms.assetid: 83e61517-6255-252c-3fee-1a0415e8b741
ms.date: 06/08/2017
---


# Trendline.NameIsAuto Property (Word)

 **True** if Microsoft Word automatically determines the name of the trendline. Read/write **Boolean** .


## Syntax

 _expression_ . **NameIsAuto**

 _expression_ A variable that represents a **[Trendline](trendline-object-word.md)** object.


## Example

The following example sets Microsoft Word to automatically determine the name for trendline one of the first chart in the active document. You should run the example on a 2-D column chart that contains a single series that has a trendline.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1) _ 
 .Trendlines(1).NameIsAuto = True 
 End If 
End With 

```


## See also


#### Concepts


[Trendline Object](trendline-object-word.md)

