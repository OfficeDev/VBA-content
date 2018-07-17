---
title: Hyperlink.Follow Method (Visio)
keywords: vis_sdr.chm15016290
f1_keywords:
- vis_sdr.chm15016290
ms.prod: visio
api_name:
- Visio.Hyperlink.Follow
ms.assetid: e415caa8-68b9-5c96-71f0-599655dc6cf3
ms.date: 06/08/2017
---


# Hyperlink.Follow Method (Visio)

Causes Microsoft Visio to navigate to a hyperlink.


## Syntax

 _expression_ . **Follow**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

Nothing


## Example

The following example draws a rectangle shape, adds a  **Hyperlink** object to the shape, sets its **Address** and **NewWindow** properties, and then uses the **Follow** method to navigate the hyperlink.

Before running this code, replace  _address_ with a valid Internet or intranet address.




```vb
 
Public Sub Follow_Example() 
 
 Dim vsoHyperlink As Visio.Hyperlink 
 
 Set vsoHyperlink = ActivePage.DrawRectangle(0,0,5,5).AddHyperlink 
 
 vsoHyperlink.Address = "http://address /" 
 vsoHyperlink.NewWindow = False 
 vsoHyperlink.Follow 
 
End Sub
```


