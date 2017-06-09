---
title: WebCommandButton.ButtonType Property (Publisher)
keywords: vbapb10.chm3932178
f1_keywords:
- vbapb10.chm3932178
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.ButtonType
ms.assetid: 9ccec0bc-4f0a-9851-0066-05ee1f144c5c
ms.date: 06/08/2017
---


# WebCommandButton.ButtonType Property (Publisher)

Returns or sets a  **PbCommandButtonType** constant that indicates whether a Web command button will clear or submit form data. Read/write.


## Syntax

 _expression_. **ButtonType**

 _expression_A variable that represents a  **WebCommandButton** object.


### Return Value

PbCommandButtonType


## Remarks

The  **ButtonType** property value can be one of the **[PbCommandButtonType](pbcommandbuttontype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example creates a new Web command submit button, assigns text to appear on its face, and specifies an e-mail address to which to send the form data.


```vb
Sub NewWebForm() 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddWebControl(Type:=pbWebControlCommandButton, _ 
 Left:=72, Top:=72, Width:=72, Height:=36) 
 With .WebCommandButton 
 .ButtonType = pbCommandButtonSubmit 
 .ButtonText = "Send Form:" 
 .EmailAddress = "someone@example.com" 
 End With 
 End With 
 End With 
End Sub
```


