---
title: Document.ClickAndTypeParagraphStyle Property (Word)
keywords: vbawd10.chm158007624
f1_keywords:
- vbawd10.chm158007624
ms.prod: word
api_name:
- Word.Document.ClickAndTypeParagraphStyle
ms.assetid: e53d3740-265f-b3ed-350a-24dd97d9f7ab
ms.date: 06/08/2017
---


# Document.ClickAndTypeParagraphStyle Property (Word)

Returns or sets the default paragraph style applied to text by the Click and Type feature in the specified document. Read/write  **Variant** .


## Syntax

 _expression_ . **ClickAndTypeParagraphStyle**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

To set the  **ClickAndTypeParagraphStyle** property, specify either the local name of the style, an integer, or a **WdBuiltinStyle** constant, or an object that represents the style. For a list of the **WdBuiltinStyle** constants, consult the Microsoft Visual Basic Object Browser.

If the  **[InUse](style-inuse-property-word.md)** property for the specified style is set to **False** , an error occurs.


## Example

This example sets the default paragraph style applied by Click and Type to Plain Text.


```vb
With ActiveDocument 
 x = "Plain Text" 
 If .Styles(x).InUse Then 
 .ClickAndTypeParagraphStyle = x 
 Else 
 MsgBox "Sorry, this style is not in use yet." 
 End If 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

