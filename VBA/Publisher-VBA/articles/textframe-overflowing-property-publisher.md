---
title: TextFrame.Overflowing Property (Publisher)
keywords: vbapb10.chm3866649
f1_keywords:
- vbapb10.chm3866649
ms.prod: publisher
api_name:
- Publisher.TextFrame.Overflowing
ms.assetid: 5a0f053b-519a-1637-0d73-992c56cdd7f0
ms.date: 06/08/2017
---


# TextFrame.Overflowing Property (Publisher)

Indicates whether the text frame contains more text than can fit into the text frame. Read-only.


## Syntax

 _expression_. **Overflowing**

 _expression_A variable that represents an  **TextFrame** object.


### Return Value

MsoTriState


## Remarks

The  **Overflowing** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|
|:-----|
| **msoFalse**|
| **msoTrue**|

## Example

This example increases the height of the selected text frame if it contains overflowing text.


```vb
Sub IncreaseTextBoxHeight() 
 With Selection.ShapeRange.TextFrame 
 If .Overflowing = msoTrue Then 
 Do 
 .Parent.Height = .Parent.Height + 36 
 Loop Until .Overflowing = msoFalse 
 End If 
 End With 
End Sub
```


