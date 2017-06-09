---
title: TextFrame2.DeleteText Method (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2.DeleteText
ms.assetid: e96a305c-085a-d807-1336-9dcc22760a7e
ms.date: 06/08/2017
---


# TextFrame2.DeleteText Method (Excel)

Deletes the text from a text frame and all the associated text properties.


## Syntax

 _expression_ . **DeleteText**

 _expression_ A variable that represents a **TextFrame2** object.


## Remarks

The associated text properties include  **Font** attributes such as bold, underline, and so on.


## Example

This example deletes the text in the text frame, if the text frame contains text.


```vb
With ActiveSheet.Shapes(1).TextFrame2 
 If .HasText Then 
 .DeleteText ()
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-excel.md)

