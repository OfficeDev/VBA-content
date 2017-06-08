---
title: WebTextBox.EchoAsterisks Property (Publisher)
keywords: vbapb10.chm4194308
f1_keywords:
- vbapb10.chm4194308
ms.prod: publisher
api_name:
- Publisher.WebTextBox.EchoAsterisks
ms.assetid: eefab42f-9fe7-e77e-50cd-c4b1b35548f1
ms.date: 06/08/2017
---


# WebTextBox.EchoAsterisks Property (Publisher)

Specifies whether asterisks should be displayed in place of text that is entered into a Web text box control. Read/write.


## Syntax

 _expression_. **EchoAsterisks**

 _expression_A variable that represents an  **WebTextBox** object.


### Return Value

MsoTrue


## Remarks

The  **EchoAsterisks** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| Displays the text entered into a Web text box control.|
| **msoTrue**| Displays asterisks in place of text entered into a Web text box control.|

## Example

This example creates a Web text box control, sets the maximum limit as ten characters, specifies that entry is required, and masks the entry with asterisks when a user enters into the control.


```vb
Sub AddPasswordTextBox() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlSingleLineTextBox, Left:=100, _ 
 Top:=100, Width:=72, Height:=15) 
 .Name = "Password" 
 With .WebTextBox 
 .Limit = 10 
 .EchoAsterisks = msoTrue 
 .RequiredControl = msoTrue 
 End With 
 End With 
End Sub
```


