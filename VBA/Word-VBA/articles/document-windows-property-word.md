---
title: Document.Windows Property (Word)
keywords: vbawd10.chm158007330
f1_keywords:
- vbawd10.chm158007330
ms.prod: word
api_name:
- Word.Document.Windows
ms.assetid: bb075fd7-2dae-18c9-f49a-0c478d840b76
ms.date: 06/08/2017
---


# Document.Windows Property (Word)

Returns a  **[Windows](windows-object-word.md)** collection that represents all windows for the specified document. Read-only.


## Syntax

 _expression_ . **Windows**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of windows for the active document, both before and after the  **NewWindow** method is run.


```vb
MsgBox Prompt:= ActiveDocument.Windows.Count &; " window(s)", _ 
 Title:= ActiveDocument.Name 
ActiveDocument.ActiveWindow.NewWindow 
MsgBox Prompt:= ActiveDocument.Windows.Count &; " windows", _ 
 Title:= ActiveDocument.Name
```


## See also


#### Concepts


[Document Object](document-object-word.md)

