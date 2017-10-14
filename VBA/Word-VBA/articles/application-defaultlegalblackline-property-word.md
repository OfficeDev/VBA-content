---
title: Application.DefaultLegalBlackline Property (Word)
keywords: vbawd10.chm158335435
f1_keywords:
- vbawd10.chm158335435
ms.prod: word
api_name:
- Word.Application.DefaultLegalBlackline
ms.assetid: a22afc29-1f7d-73af-75c2-7ce2fbe2250f
ms.date: 06/08/2017
---


# Application.DefaultLegalBlackline Property (Word)

 **True** for Microsoft Word to compare and merge documents using the **Legal blackline** option in the **Compare and Merge Documents** dialog box. Read/write **Boolean** .


## Syntax

 _expression_ . **DefaultLegalBlackline**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example enables Word's Legal blackline option for comparing and merging legal documents.


```vb
Sub CreateLegalBlackline() 
 Application.DefaultLegalBlackline = True 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

