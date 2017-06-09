---
title: Version.Comment Property (Word)
keywords: vbawd10.chm162792428
f1_keywords:
- vbawd10.chm162792428
ms.prod: word
api_name:
- Word.Version.Comment
ms.assetid: 52ca9077-0295-3059-b699-6fa97ad45991
ms.date: 06/08/2017
---


# Version.Comment Property (Word)

Returns the comment associated with the specified version of a document. Read-only  **String** .


## Syntax

 _expression_ . **Comment**

 _expression_ A variable that represents a **[Version](version-object-word.md)** object.


## Example

This example displays the comment text for the first version of the active document.


```vb
If ActiveDocument.Versions.Count >= 1 Then 
 MsgBox Prompt:=ActiveDocument.Versions(1).Comment, _ 
 Title:="First Version Comment" 
End If
```

This example saves a version of the document with the user's comment and then displays the comment.




```vb
Dim verTemp As Versions 
Dim strComment As String 
Dim lngCount As Long 
 
Set verTemp = ActiveDocument.Versions 
 
strComment = InputBox("Type a comment") 
verTemp.Save Comment:=strComment 
lngCount = verTemp.Count 
MsgBox Prompt:=verTemp(lngCount).Comment, _ 
 Title:=verTemp(lngCount).SavedBy
```


## See also


#### Concepts


[Version Object](version-object-word.md)

