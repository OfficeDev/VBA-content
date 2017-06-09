---
title: Revision.FormatDescription Property (Word)
keywords: vbawd10.chm159449097
f1_keywords:
- vbawd10.chm159449097
ms.prod: word
api_name:
- Word.Revision.FormatDescription
ms.assetid: 5178a4d2-ae38-a0e7-4df4-3bac2789d37d
ms.date: 06/08/2017
---


# Revision.FormatDescription Property (Word)

Returns a  **String** representing a description of tracked formatting changes in a revision. Read-only.


## Syntax

 _expression_ . **FormatDescription**

 _expression_ An expression that returns a **[Revision](revision-object-word.md)** object.


## Example

This example displays a description for each of the formatting changes made in a document with tracked changes.


```vb
Sub FmtChanges() 
 Dim revFmtRev As Revision 
 
 For Each revFmtRev In ActiveDocument.Revisions 
 If revFmtRev.FormatDescription <> "" Then 
 MsgBox "Format changes made : " &; revFmtRev.FormatDescription 
 End If 
 Next 
End Sub
```


## See also


#### Concepts


[Revision Object](revision-object-word.md)

