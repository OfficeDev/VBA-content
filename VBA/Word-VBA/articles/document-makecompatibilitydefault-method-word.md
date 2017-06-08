---
title: Document.MakeCompatibilityDefault Method (Word)
keywords: vbawd10.chm158007415
f1_keywords:
- vbawd10.chm158007415
ms.prod: word
api_name:
- Word.Document.MakeCompatibilityDefault
ms.assetid: 06c3cede-312c-aacf-3780-4d79dd7c6fc3
ms.date: 06/08/2017
---


# Document.MakeCompatibilityDefault Method (Word)

Sets the compatibility options.


## Syntax

 _expression_ . **MakeCompatibilityDefault**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The compatibility options are located on the  **Compatibility** tab in the **Options** dialog box as the default settings for new documents.


## Example

This example sets a few compatibility options for the active document and then makes the current compatibility options the default settings.


```vb
With ActiveDocument 
 .Compatibility(wdSuppressSpBfAfterPgBrk) = True 
 .Compatibility(wdExpandShiftReturn) = True 
 .Compatibility(wdUsePrinterMetrics) = True 
 .Compatibility(wdNoLeading) = False 
 .MakeCompatibilityDefault 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

