---
title: WebTextBox.RequiredControl Property (Publisher)
keywords: vbapb10.chm4194310
f1_keywords:
- vbapb10.chm4194310
ms.prod: publisher
api_name:
- Publisher.WebTextBox.RequiredControl
ms.assetid: 32e18d4b-7af0-b079-4baf-9acc07c3c37d
ms.date: 06/08/2017
---


# WebTextBox.RequiredControl Property (Publisher)

Specifies whether an entry into a Web text box control is required. Read/write.


## Syntax

 _expression_. **RequiredControl**

 _expression_A variable that represents a  **WebTextBox** object.


### Return Value

MsoTriState


## Remarks

The  **RequiredControl** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|Indicates entry into the specified Web text box control is not required.|
| **msoTrue**| Indicates entry into the specified Web text box control is required.|

## Example

This example creates a new Web text box control in the active publication, sets the default text and the character limit for the text box, and specifies that an entry is required.


```vb
Sub AddWebTextBoxControl() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlMultiLineTextBox, Left:=72, _ 
 Top:=72, Width:=300, Height:=100).WebTextBox 
 .DefaultText = "Please enter text here." 
 .Limit = 200 
 .RequiredControl = msoTrue 
 End With 
End Sub
```


