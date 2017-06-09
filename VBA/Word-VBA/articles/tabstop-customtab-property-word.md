---
title: TabStop.CustomTab Property (Word)
keywords: vbawd10.chm156500071
f1_keywords:
- vbawd10.chm156500071
ms.prod: word
api_name:
- Word.TabStop.CustomTab
ms.assetid: c909f223-7e5d-6a2b-317f-12f735e43921
ms.date: 06/08/2017
---


# TabStop.CustomTab Property (Word)

 **True** if the specified tab stop is a custom tab stop. Read-only **Boolean** .


## Syntax

 _expression_ . **CustomTab**

 _expression_ A variable that represents a **[TabStop](tabstop-object-word.md)** object.


## Example

This example cycles through the collection of tab stops in the first paragraph in the active document, and left-aligns any custom tab stops that it finds.


```vb
Dim tsLoop As TabStop 
 
For each tsLoop in ActiveDocument.Paragraphs(1).TabStops 
 If tsLoop.CustomTab = True Then 
 tsLoop.Alignment = wdAlignTabLeft 
 End If 
Next tsLoop
```


## See also


#### Concepts


[TabStop Object](tabstop-object-word.md)

