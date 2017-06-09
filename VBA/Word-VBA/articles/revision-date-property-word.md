---
title: Revision.Date Property (Word)
keywords: vbawd10.chm159449090
f1_keywords:
- vbawd10.chm159449090
ms.prod: word
api_name:
- Word.Revision.Date
ms.assetid: 3c8941e1-7b1e-23d0-89f6-a83db6c00f20
ms.date: 06/08/2017
---


# Revision.Date Property (Word)

The date and time that the tracked change was made. Read-only  **Date** .


## Syntax

 _expression_ . **Date**

 _expression_ A variable that represents a **[Revision](revision-object-word.md)** object.


## Example

This example displays the date and time of the next tracked change found in the active document.


```vb
Dim revTemp As Revision 
 
If ActiveDocument.Revisions.Count >= 1 Then 
 Set revTemp = Selection.NextRevision 
 If Not (revTemp Is Nothing) Then MsgBox revTemp.Date 
End If
```


## See also


#### Concepts


[Revision Object](revision-object-word.md)

