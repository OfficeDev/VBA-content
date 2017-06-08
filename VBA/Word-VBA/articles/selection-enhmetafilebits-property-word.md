---
title: Selection.EnhMetaFileBits Property (Word)
keywords: vbawd10.chm158662971
f1_keywords:
- vbawd10.chm158662971
ms.prod: word
api_name:
- Word.Selection.EnhMetaFileBits
ms.assetid: ecc28cc8-6c0f-3207-f52c-4a7b77c23445
ms.date: 06/08/2017
---


# Selection.EnhMetaFileBits Property (Word)

Returns a  **Variant** that represents a picture representation of how a selection or range of text appears.


## Syntax

 _expression_ . **EnhMetaFileBits**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

The  **EnhMetaFileBits** property returns an array of bytes, which can be used with the Microsoft Windows 32 Application Programming Interface from within the Microsoft Visual Basic or Microsoft C++ development environment.


## Example

The following example returns the  **EnhMetaFileBits** property.


```vb
Dim bytSelection() As Byte 
 
bytSelection = Selection.EnhMetaFileBits
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

