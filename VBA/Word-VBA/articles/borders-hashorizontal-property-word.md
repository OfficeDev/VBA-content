---
title: Borders.HasHorizontal Property (Word)
keywords: vbawd10.chm154927131
f1_keywords:
- vbawd10.chm154927131
ms.prod: word
api_name:
- Word.Borders.HasHorizontal
ms.assetid: 5a5863c8-8f0d-67f9-6e1f-2a4dd6b4fbc6
ms.date: 06/08/2017
---


# Borders.HasHorizontal Property (Word)

 **True** if a horizontal border can be applied to the object. Read-only **Boolean** .


## Syntax

 _expression_ . **HasHorizontal**

 _expression_ A variable that represents a **[Borders](borders-object-word.md)** object.


## Remarks

Horizontal borders can be applied to ranges that contain cells in two or more rows of a table or ranges that contain two or more paragraphs.


## Example

This example applies single-line horizontal borders, if the selection supports horizontal borders.


```vb
If Selection.Borders.HasHorizontal = True Then 
 Selection.Borders(wdBorderHorizontal).LineStyle = _ 
 wdLineStyleSingle 
End If
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

