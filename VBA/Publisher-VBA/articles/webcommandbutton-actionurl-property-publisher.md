---
title: WebCommandButton.ActionURL Property (Publisher)
keywords: vbapb10.chm3932163
f1_keywords:
- vbapb10.chm3932163
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.ActionURL
ms.assetid: ede9b18f-1be1-9572-9b78-7dbe0817cfe7
ms.date: 06/08/2017
---


# WebCommandButton.ActionURL Property (Publisher)

Returns or sets a  **String** that represents the URL of the server-side script to execute in response to a Submit button click. Read/write.


## Syntax

 _expression_. **ActionURL**

 _expression_A variable that represents a  **WebCommandButton** object.


### Return Value

String


## Remarks

The default value for the  **ActionURL** property is "http://example.microsoft.com/~user/ispscript.cgi". This property is ignored for Reset command buttons.


## Example

This example creates a Web form Submit command button and sets the script path and file name to run when a user clicks the button.


```vb
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "http://www.tailspintoys.com/" &; _ 
 "scripts/ispscript.cgi" 
 End With 
End Sub
```


