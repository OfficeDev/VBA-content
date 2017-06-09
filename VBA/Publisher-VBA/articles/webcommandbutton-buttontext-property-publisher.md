---
title: WebCommandButton.ButtonText Property (Publisher)
keywords: vbapb10.chm3932164
f1_keywords:
- vbapb10.chm3932164
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.ButtonText
ms.assetid: 0a9a7bd9-de7e-7e80-0aa2-7cefda17f354
ms.date: 06/08/2017
---


# WebCommandButton.ButtonText Property (Publisher)

Returns or sets a  **String** that represents the text that appears on the face of a Web command button. Read/write.


## Syntax

 _expression_. **ButtonText**

 _expression_A variable that represents a  **WebCommandButton** object.


### Return Value

String


## Example

This example creates a new Web command button, assigns text to appear on its face, and specifies an e-mail address to which to send the form data.


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


