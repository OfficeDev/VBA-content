---
title: Revision.Reject Method (Word)
keywords: vbawd10.chm159449190
f1_keywords:
- vbawd10.chm159449190
ms.prod: word
api_name:
- Word.Revision.Reject
ms.assetid: e97603c6-2310-ad82-7145-66a640a05c04
ms.date: 06/08/2017
---


# Revision.Reject Method (Word)

Rejects the specified tracked change. The revision marks are removed, leaving the original text intact.


## Syntax

 _expression_ . **Reject**

 _expression_ Required. A variable that represents a **[Revision](revision-object-word.md)** object.


## Remarks

Formatting changes cannot be rejected.


## Example

This example rejects the next tracked change found in the active document.


```vb
Dim revNext As Revision 
 
If ActiveDocument.Revisions.Count >= 1 Then 
 Set revNext = Selection.NextRevision 
 If Not (revNext Is Nothing) Then revNext.Reject 
End If
```

This example rejects the tracked changes in the first paragraph.




```vb
Dim rngTemp As Range 
Dim revLoop As Revision 
 
Set rngTemp = ActiveDocument.Paragraphs(1).Range 
For Each revLoop In rngTemp.Revisions 
 revLoop.Reject 
Next revLoop
```

This example rejects the first tracked change in the selection.




```vb
Dim rngTemp As Range 
 
Set rngTemp = Selection.Range 
If rngTemp.Revisions.Count >= 1 Then _ 
 rngTemp.Revisions(1).Reject
```


## See also


#### Concepts


[Revision Object](revision-object-word.md)

