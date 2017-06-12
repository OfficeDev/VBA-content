---
title: Subdocuments.Expanded Property (Word)
keywords: vbawd10.chm159907842
f1_keywords:
- vbawd10.chm159907842
ms.prod: word
api_name:
- Word.Subdocuments.Expanded
ms.assetid: 99879e46-d762-64e8-fa07-c88f3dceb3eb
ms.date: 06/08/2017
---


# Subdocuments.Expanded Property (Word)

 **True** if the subdocuments in the specified document are expanded. Read/write **Boolean** .


## Syntax

 _expression_ . **Expanded**

 _expression_ A variable that represents a **[Subdocument](subdocument-object-word.md)** object.


## Example

This example expands all subdocuments in the active master document.


```vb
If ActiveDocument.Subdocuments.Count >= 1 Then 
 ActiveDocument.Subdocuments.Expanded = True 
End If
```

This example switches the  **Expanded** property between expanding all subdocuments in the active window and collapsing all subdocuments in the active document.




```vb
ActiveDocument.Subdocuments.Expanded = _ 
 Not ActiveDocument.Subdocuments.Expanded
```

This example determines whether the subdocuments in Report.doc are expanded and then displays a message indicating their status.




```vb
If Documents("Report.doc").Subdocuments.Expanded = True Then 
 MsgBox "All available information is displayed." 
Else 
 MsgBox "Expand subdocuments for more information." 
End If
```


## See also


#### Concepts


[Subdocuments Collection Object](subdocuments-object-word.md)

