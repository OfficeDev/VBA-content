---
title: Document.XMLShowAdvancedErrors Property (Word)
keywords: vbawd10.chm158007774
f1_keywords:
- vbawd10.chm158007774
ms.prod: word
api_name:
- Word.Document.XMLShowAdvancedErrors
ms.assetid: 56ddb6ee-f2fd-fa8e-5f07-a5af4d749652
ms.date: 06/08/2017
---


# Document.XMLShowAdvancedErrors Property (Word)

Returns or sets a  **Boolean** that represents whether error message text is generated from the built-in Microsoft Word error messages or from the Microsoft XML Core Services (MSXML) 5.0 component included with Office.


## Syntax

 _expression_ . **XMLShowAdvancedErrors**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

Using advanced error messages from the MSXML 5.0 component provides more specific error messages. There are approximately 500 error messages provided through the XML Core Services that are accessible when the  **XMLShowAdvancedErrors** property is **True** .

Word has approximately 50 built-in generic schema errors. When the  **XMLShowAdvancedErrors** property is **False** , Word uses the built-in error messages for errors generated in XML documents.


## Example

The following example enables advanced error messages in the active document.


```vb
ActiveDocument.XMLShowAdvancedErrors = True
```


## See also


#### Concepts


[Document Object](document-object-word.md)

