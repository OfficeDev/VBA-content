---
title: WebCommandButton.HiddenFields Property (Publisher)
keywords: vbapb10.chm3932177
f1_keywords:
- vbapb10.chm3932177
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.HiddenFields
ms.assetid: 187553fb-a4d3-a1fb-f583-49e1d76992ec
ms.date: 06/08/2017
---


# WebCommandButton.HiddenFields Property (Publisher)

Returns a  **WebHiddenFields** object that represents hidden Web fields attached to a Submit command button.


## Syntax

 _expression_. **HiddenFields**

 _expression_A variable that represents a  **WebCommandButton** object.


### Return Value

WebHiddenFields


## Example

This example adds a new hidden Web field to a new Submit command button.


```vb
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 End With 
 .Item(1).WebCommandButton.HiddenFields.Add _ 
 Name:="User", Value:="PowerUser" 
 End With 
End Sub
```


