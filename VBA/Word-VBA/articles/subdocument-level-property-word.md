---
title: Subdocument.Level Property (Word)
keywords: vbawd10.chm159973382
f1_keywords:
- vbawd10.chm159973382
ms.prod: word
api_name:
- Word.Subdocument.Level
ms.assetid: 5a4d20aa-8801-77b7-ad86-6c0e26179bef
ms.date: 06/08/2017
---


# Subdocument.Level Property (Word)

Returns the heading level used to create the subdocument. Read-only  **Long** .


## Syntax

 _expression_ . **Level**

 _expression_ Required. A variable that represents a **[Subdocument](subdocument-object-word.md)** object.


## Example

This example looks through each subdocument in the active document and displays the subdocument's heading level.


```vb
i = 1 
If ActiveDocument.Subdocuments.Count > = 1 Then 
 For each s in ActiveDocument.Subdocuments 
 MsgBox "The heading level for SubDoc " &; i _ 
 &; " is " &; s.Level 
 i = i + 1 
 Next s 
Else 
 MsgBox "There are no subdocuments defined." 
End If
```


## See also


#### Concepts


[Subdocument Object](subdocument-object-word.md)

