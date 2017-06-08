---
title: Document.SaveFormat Property (Word)
keywords: vbawd10.chm158007355
f1_keywords:
- vbawd10.chm158007355
ms.prod: word
api_name:
- Word.Document.SaveFormat
ms.assetid: f8d31365-1935-307f-3663-d6e769944489
ms.date: 06/08/2017
---


# Document.SaveFormat Property (Word)

Returns the file format of the specified document or file converter. Read-only  **Long** .


## Syntax

 _expression_ . **SaveFormat**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **SaveFormat** property will be a unique number that specifies an external file converter or a **WdSaveFormat** constant.

Use the value of the  **SaveFormat** property for the _FileFormat_ argument of the **[SaveAs2](document-saveas2-method-word.md)** method to save a document in a file format for which there isn't a corresponding **WdSaveFormat** constant.


## Example

If the active document is a Rich Text Format (RTF) document, this example saves it as a Microsoft Word document.


```vb
If ActiveDocument.SaveFormat = wdFormatRTF Then 
 ActiveDocument.SaveAs FileFormat:=wdFormatDocument 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

