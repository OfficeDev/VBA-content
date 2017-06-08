---
title: Document.Path Property (Word)
keywords: vbawd10.chm158007299
f1_keywords:
- vbawd10.chm158007299
ms.prod: word
api_name:
- Word.Document.Path
ms.assetid: 809b41fb-c410-5bcb-c808-780ad5232e6f
ms.date: 06/08/2017
---


# Document.Path Property (Word)

Returns the disk or Web path to the document. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "http://MyServer". Use the  **PathSeparator** property to add the character that separates folders and drive letters. Use the **Name** property to return the file name without the path and use the **FullName** property to return the file name and the path together.


 **Note**  You can use the  **PathSeparator** property to build Web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## Example

This example displays the path and file name of the active document.


```vb
MsgBox ActiveDocument.Path &; Application.PathSeparator &; _ 
 ActiveDocument.Name
```

This example changes the current folder to the path of the template attached to the active document.




```
ChDir ActiveDocument.AttachedTemplate.Path
```


## See also


#### Concepts


[Document Object](document-object-word.md)

