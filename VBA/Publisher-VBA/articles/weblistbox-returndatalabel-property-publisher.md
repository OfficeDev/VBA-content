---
title: WebListBox.ReturnDataLabel Property (Publisher)
keywords: vbapb10.chm4063237
f1_keywords:
- vbapb10.chm4063237
ms.prod: publisher
api_name:
- Publisher.WebListBox.ReturnDataLabel
ms.assetid: 0c9a6942-1cc7-92b6-116e-836e79560084
ms.date: 06/08/2017
---


# WebListBox.ReturnDataLabel Property (Publisher)

Returns or sets a  **String** that represents the text used by the Web page to label the specified Web object when the page is submitted. Read/write.


## Syntax

 _expression_. **ReturnDataLabel**

 _expression_A variable that represents a  **WebListBox** object.


## Example

This example creates a new Web text box and specifies the label for the text in the text box when the page is submitted.


```vb
Sub LabelWebTextBoxControl() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=300, Height:=15).WebTextBox 
 .DefaultText = "Please enter your name here" 
 .Limit = 70 
 .RequiredControl = msoTrue 
 .ReturnDataLabel = "Full_Name" 
 End With 
End Sub
```


