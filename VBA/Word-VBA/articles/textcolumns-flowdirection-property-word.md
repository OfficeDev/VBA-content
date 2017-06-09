---
title: TextColumns.FlowDirection Property (Word)
keywords: vbawd10.chm158531688
f1_keywords:
- vbawd10.chm158531688
ms.prod: word
api_name:
- Word.TextColumns.FlowDirection
ms.assetid: 65d2791e-f3ae-a3df-5d93-959750516b11
ms.date: 06/08/2017
---


# TextColumns.FlowDirection Property (Word)

Returns or sets the direction in which text flows from one text column to the next. Read/write  **WdFlowDirection** .


## Syntax

 _expression_ . **FlowDirection**

 _expression_ Required. A variable that represents a **[TextColumns](textcolumns-objectword.md)** collection.


## Example

This example sets the flow direction so that text flows through the specified columns from right to left.


```vb
ActiveDocument.PageSetup.TextColumns.FlowDirection = _ 
 wdFlowRtl
```


## See also


#### Concepts


[TextColumns Collection Object](textcolumns-objectword.md)

