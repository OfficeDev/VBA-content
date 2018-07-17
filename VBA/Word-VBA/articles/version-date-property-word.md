---
title: Version.Date Property (Word)
keywords: vbawd10.chm162792429
f1_keywords:
- vbawd10.chm162792429
ms.prod: word
api_name:
- Word.Version.Date
ms.assetid: c46596fc-e6ff-4158-ba83-d83a6e84400b
ms.date: 06/08/2017
---


# Version.Date Property (Word)

The date and time that the document version was saved. Read-only  **Date** .


## Syntax

 _expression_ . **Date**

 _expression_ A variable that represents a **[Version](version-object-word.md)** object.


## Example

This example displays the date and time that the last version of the active document was saved.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 

```


```vb
If docActive.Path <> "" Then MsgBox _ 
 docActive.Versions(docActive.Versions.Count).Date
```

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


[Version Object](version-object-word.md)

