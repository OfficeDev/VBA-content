---
title: ChartGroup.ShowNegativeBubbles Property (Word)
keywords: vbawd10.chm263454758
f1_keywords:
- vbawd10.chm263454758
ms.prod: word
api_name:
- Word.ChartGroup.ShowNegativeBubbles
ms.assetid: 6102a2dd-2ef8-2047-f14a-ca64e88a0565
ms.date: 06/08/2017
---


# ChartGroup.ShowNegativeBubbles Property (Word)

 **True** if negative bubbles are shown for the chart group. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowNegativeBubbles**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

The property is valid only for bubble charts. 


## Example

The following example causes Microsoft Word to display negative bubbles for the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).ShowNegativeBubbles = True 
 End If 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

