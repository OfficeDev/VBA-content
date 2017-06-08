---
title: InvisibleApp.Window Property (Visio)
keywords: vis_sdr.chm17551480
f1_keywords:
- vis_sdr.chm17551480
ms.prod: visio
api_name:
- Visio.InvisibleApp.Window
ms.assetid: 6b693eb6-51c0-8bc7-69d4-f5f4fc921d68
ms.date: 06/08/2017
---


# InvisibleApp.Window Property (Visio)

Returns the window associated with the current instance of Microsoft Visio. Read-only.


## Syntax

 _expression_ . **Window**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Window


## Example

The following macro shows how to use the  **Window** property to print the caption of the window associated with the current instance of Visio in the Immediate window.


```vb
 
Public Sub Window_Example() 
 
 Debug.Print Application.Window.Caption 
 
End Sub
```


