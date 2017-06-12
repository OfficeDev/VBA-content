---
title: TabStop.Leader Property (Word)
keywords: vbawd10.chm156500069
f1_keywords:
- vbawd10.chm156500069
ms.prod: word
api_name:
- Word.TabStop.Leader
ms.assetid: 3e483648-b48f-c8e0-93c0-e83771c48299
ms.date: 06/08/2017
---


# TabStop.Leader Property (Word)

Returns or sets the leader for the specified  **TabStop** object. Read/write **WdTabLeader** .


## Syntax

 _expression_ . **Leader**

 _expression_ Required. A variable that represents a **[TabStop](tabstop-object-word.md)** object.


## Example

This example changes the leader for all tab stops that have a leader to dashes for all the paragraphs in the active document.


```vb
Dim tsLoop As TabStop 
 
For each tsLoop in ActiveDocument.Paragraphs.TabStops 
 If tsLoop.Leader <> wdTabLeaderSpaces Then 
 tsLoop.Leader = wdTabLeaderDashes 
 End If 
Next tsLoop
```


## See also


#### Concepts


[TabStop Object](tabstop-object-word.md)

