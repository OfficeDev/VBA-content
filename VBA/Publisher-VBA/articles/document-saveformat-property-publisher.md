---
title: Document.SaveFormat Property (Publisher)
keywords: vbapb10.chm196656
f1_keywords:
- vbapb10.chm196656
ms.prod: publisher
api_name:
- Publisher.Document.SaveFormat
ms.assetid: 545f0411-899f-ffe3-e844-8c2922a357f0
ms.date: 06/08/2017
---


# Document.SaveFormat Property (Publisher)

Indicates the file format of the specified document. Read-only.


## Syntax

 _expression_. **SaveFormat**

 _expression_A variable that represents a  **Document** object.


### Return Value

PbFileFormat


## Remarks

The  **SaveFormat** property value can be one of the **[PbFileFormat](pbfileformat-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

If the active publication is in the Publisher 2000 format, this example saves it in Rich Text Format (RTF).


```vb
Sub SaveAsRTF() 
 
 If Application.ActiveDocument.SaveFormat = pbFilePublisher2000 Then 
 ActiveDocument.SaveAs "Flyer3", pbFileRTF 
 End If 
 
End Sub
```


