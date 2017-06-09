---
title: TextFrame.VerticalTextAlignment Property (Publisher)
keywords: vbapb10.chm3866660
f1_keywords:
- vbapb10.chm3866660
ms.prod: publisher
api_name:
- Publisher.TextFrame.VerticalTextAlignment
ms.assetid: cd809f00-b092-c483-fe99-2aa8043fb684
ms.date: 06/08/2017
---


# TextFrame.VerticalTextAlignment Property (Publisher)

Returns or sets a  **PbVerticalTextAlignmentType**constant that represents the vertical alignment of text in a text box. Read/write.


## Syntax

 _expression_. **VerticalTextAlignment**

 _expression_A variable that represents a  **TextFrame** object.


## Remarks

The  **VerticalTextAlignment** property value can be one of these **PbVerticalTextAlignmentType** constants.



| **pbVerticalTextAlignmentBottom**|
| **pbVerticalTextAlignmentCenter**|
| **pbVerticalTextAlignmentTop**|

## Example

This example vertically centers the text in the specified text frame. This example assumes there is at least one shape on the first page of the active publication.


```vb
Sub SetVerticalAlignment() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .VerticalTextAlignment = pbVerticalTextAlignmentCenter 
End Sub
```


