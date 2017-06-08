---
title: Selection.CopyAsPicture Method (Word)
keywords: vbawd10.chm158662823
f1_keywords:
- vbawd10.chm158662823
ms.prod: word
api_name:
- Word.Selection.CopyAsPicture
ms.assetid: f5c73e30-1601-62a7-ec0e-2dc49c6f51fe
ms.date: 06/08/2017
---


# Selection.CopyAsPicture Method (Word)

The  **CopyAsPicture** method works the same way as the **Copy** method.


## Syntax

 _expression_ . **CopyAsPicture**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example copies the contents of the active document as a picture and pastes it as a picture at the end of the document.


```vb
Sub CopyPasteAsPicture() 
 ActiveDocument.Content.Select 
 With Selection 
 .CopyAsPicture 
 .Collapse Direction:=wdCollapseEnd 
 .PasteSpecial DataType:=wdPasteMetafilePicture 
 End With 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

