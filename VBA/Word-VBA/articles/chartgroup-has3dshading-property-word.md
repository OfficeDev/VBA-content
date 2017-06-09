---
title: ChartGroup.Has3DShading Property (Word)
keywords: vbawd10.chm263454766
f1_keywords:
- vbawd10.chm263454766
ms.prod: word
api_name:
- Word.ChartGroup.Has3DShading
ms.assetid: 095f5bc7-86aa-2c09-c52c-6e6d5a4deb16
ms.date: 06/08/2017
---


# ChartGroup.Has3DShading Property (Word)

 **True** if a chart group has three-dimensional shading. Read/write **Boolean** .


## Syntax

 _expression_ . **Has3DShading**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

Setting  **Has3DShading** to **False** removes the 3-D shading effect from the chart (rendering it as flat). Setting **Has3DShading** to **True** sets the chart content to the default.


## Example

The following example adds 3-D shading to the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).Has3DShading = True 
 End If 
End With 

```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

