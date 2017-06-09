---
title: Window.DocumentMap Property (Word)
keywords: vbawd10.chm157417497
f1_keywords:
- vbawd10.chm157417497
ms.prod: word
api_name:
- Word.Window.DocumentMap
ms.assetid: e7f084f8-303b-d710-00fc-522eab6e3814
ms.date: 06/08/2017
---


# Window.DocumentMap Property (Word)

 **True** if the document map is visible. Read/write **Boolean** .


## Syntax

 _expression_ . **DocumentMap**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example toggles the document map for the active window.


```vb
ActiveDocument.ActiveWindow.DocumentMap = _ 
 Not ActiveDocument.ActiveWindow.DocumentMap
```

This example displays the document map in the window for Sales.doc.




```vb
Dim docSales As Document 
 
Set docSales = _ 
 Documents.Open(FileName:="C:\Documents\Sales.doc") 
 
docSales.ActiveWindow.DocumentMap = True
```


## See also


#### Concepts


[Window Object](window-object-word.md)

