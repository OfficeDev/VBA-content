---
title: OLEFormat.PreserveFormattingOnUpdate Property (Word)
keywords: vbawd10.chm154337392
f1_keywords:
- vbawd10.chm154337392
ms.prod: word
api_name:
- Word.OLEFormat.PreserveFormattingOnUpdate
ms.assetid: 2292fee8-42c6-274c-2ef8-de21af16314a
ms.date: 06/08/2017
---


# OLEFormat.PreserveFormattingOnUpdate Property (Word)

 **True** preserves formatting done in Microsoft Word to a linked OLE object, such as a table linked to a Microsoft Excel spreadsheet. Read/write **Boolean** .


## Syntax

 _expression_ . **PreserveFormattingOnUpdate**

 _expression_ A variable that represents a **[OLEFormat](oleformat-object-word.md)** object.


## Remarks

When  **PreserveFormattingOnUpdate** is set to **True** , formatting changes made to the object in Word is preserved when the object is updated. Word updates only the content in the linked object.


## Example

This example preserves the formatting of the first shape in the current document, assuming the first shape in the document is a linked OLE object.


```vb
Sub PreserveFmtg() 
 ActiveDocument.Shapes(1).OLEFormat _ 
 .PreserveFormattingOnUpdate = True 
End Sub
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

