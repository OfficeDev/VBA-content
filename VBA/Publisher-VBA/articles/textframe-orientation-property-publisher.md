---
title: TextFrame.Orientation Property (Publisher)
keywords: vbapb10.chm3866659
f1_keywords:
- vbapb10.chm3866659
ms.prod: publisher
api_name:
- Publisher.TextFrame.Orientation
ms.assetid: f510e624-6322-4054-5e7f-8688c5ea817a
ms.date: 06/08/2017
---


# TextFrame.Orientation Property (Publisher)

Returns or sets a  **PbTextOrientation**constant that represents how text flows in a text box. Read/write.


## Syntax

 _expression_. **Orientation**

 _expression_A variable that represents a  **TextFrame** object.


### Return Value

PbTextOrientation


## Remarks

The  **Orientation** property value can be one of the **[PbTextOrientation](pbtextorientation-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example sets the text orientation in the specified text box to vertical so that text flows from top to bottom. This assumes there is at least one shape on page one of the active publication.






```vb
Sub SetVerticalTextBox() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .Orientation = pbTextOrientationVerticalEastAsia 
End Sub
```


