---
title: WebTextBox.DefaultText Property (Publisher)
keywords: vbapb10.chm4194307
f1_keywords:
- vbapb10.chm4194307
ms.prod: publisher
api_name:
- Publisher.WebTextBox.DefaultText
ms.assetid: 348c1bc2-61c9-f89f-5e7a-b73ddaa3d216
ms.date: 06/08/2017
---


# WebTextBox.DefaultText Property (Publisher)

Returns or sets a  **String** that represents the default text in a Web text box control. Read/write.


## Syntax

 _expression_. **DefaultText**

 _expression_A variable that represents a  **WebTextBox** object.


### Return Value

String


## Example

This example creates a new Web text box control in the active publication, sets the default text and the character limit for the text box, and specifies that it is a required control.


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


