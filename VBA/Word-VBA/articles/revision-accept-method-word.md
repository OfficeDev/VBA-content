---
title: Revision.Accept Method (Word)
keywords: vbawd10.chm159449189
f1_keywords:
- vbawd10.chm159449189
ms.prod: word
api_name:
- Word.Revision.Accept
ms.assetid: 3e98b15a-edc3-dc85-0297-288886d8c479
ms.date: 06/08/2017
---


# Revision.Accept Method (Word)

Accepts the specified tracked change, removes the revision mark, and incorporates the change into the document.


## Syntax

 _expression_ . **Accept**

 _expression_ Required. A variable that represents a **[Revision](revision-object-word.md)** object.


## Example

This example accepts the next tracked change found if the change type is inserted text.


```vb
Set revNext = Selection.NextRevision(Wrap:=True) 
 
If Not (revNext Is Nothing) Then 
 If revNext.Type = wdRevisionInsert Then revNext.Accept 
End If
```

This example accepts all the tracked changes in the selection.




```vb
Dim revLoop As Revision 
Dim rngSelection As Range 
 
Set rngSelection = Selection.Range 
For Each revLoop In rngSelection.Revisions 
 revLoop.Accept 
Next revLoop
```


## See also


#### Concepts


[Revision Object](revision-object-word.md)

