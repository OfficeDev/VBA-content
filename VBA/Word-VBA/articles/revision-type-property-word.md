---
title: Revision.Type Property (Word)
keywords: vbawd10.chm159449092
f1_keywords:
- vbawd10.chm159449092
ms.prod: word
api_name:
- Word.Revision.Type
ms.assetid: 290549a0-5ace-7222-1e7c-5469129c8350
ms.date: 06/08/2017
---


# Revision.Type Property (Word)

Returns the revision type. Read-only  **[WdRevisionType](wdrevisiontype-enumeration-word.md)** .


## Syntax

 _expression_ . **Type**

 _expression_ Required. A variable that represents a **[Revision](revision-object-word.md)** object.


## Example

This example accepts the next revision in the active document if the revision type is inserted text.


```vb
Set myRev = Selection.NextRevision 
If Not (myRev Is Nothing) Then 
 If myRev.Type = wdRevisionInsert Then myRev.Accept 
End If
```


## See also


#### Concepts


[Revision Object](revision-object-word.md)

