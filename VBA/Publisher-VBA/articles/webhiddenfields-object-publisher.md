---
title: WebHiddenFields Object (Publisher)
keywords: vbapb10.chm4063231
f1_keywords:
- vbapb10.chm4063231
ms.prod: publisher
api_name:
- Publisher.WebHiddenFields
ms.assetid: 8ced4021-fa99-39dd-e880-b9793426871f
ms.date: 06/08/2017
---


# WebHiddenFields Object (Publisher)

Represents hidden Web fields that allow a Web page to pass non-visible data to the Web server when a Web page is submitted. The  **WebHiddenFields** object enables control of all the hidden fields attached to a Submit command button.
 


## Example

Use the  **HiddenFields** property to access hidden Web fields. This example adds a new hidden Web field to a new Submit command button.
 

 

```
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .HiddenFields.Add Name:="User", Value:="PowerUser" 
 End With 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](webhiddenfields-add-method-publisher.md)|
|[Delete](webhiddenfields-delete-method-publisher.md)|
|[Item](webhiddenfields-item-method-publisher.md)|
|[Name](webhiddenfields-name-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](webhiddenfields-application-property-publisher.md)|
|[Count](webhiddenfields-count-property-publisher.md)|
|[Parent](webhiddenfields-parent-property-publisher.md)|

