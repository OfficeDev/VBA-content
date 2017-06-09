---
title: InvisibleApp.Addons Property (Visio)
keywords: vis_sdr.chm17513060
f1_keywords:
- vis_sdr.chm17513060
ms.prod: visio
api_name:
- Visio.InvisibleApp.Addons
ms.assetid: a0f1039c-595a-18cf-1484-842069db677f
ms.date: 06/08/2017
---


# InvisibleApp.Addons Property (Visio)

Returns the  **Addons** collection of an **Application** or **InvisibleApp** object. Read-only.


## Syntax

 _expression_ . **Addons**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Addons


## Remarks

The  **Addons** collection includes an **Addon** object for each add-on in the folders specified by the **AddonPaths** property and for each add-on that is added dynamically to the collection by other add-ons.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the  **Addons** collection and add an add-on to it.

Before running this macro, replace  _path\filename_ with a valid path and file name for an add-on in your Visio project.




```vb
 
Public Sub Addons_Example() 
 
 Dim vsoAddons As Visio.Addons 
 Dim vsoAddon As Visio.Addon 
 
 'Add an add-on to the Addons collection. 
 Set vsoAddons = Visio.Addons 
 Set vsoAddon = vsoAddons.Add("path\filename ") 
 
End Sub
```


