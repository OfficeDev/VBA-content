---
title: DrawingControl.Window Property (Visio)
keywords: vis_sdr.chm51020
f1_keywords:
- vis_sdr.chm51020
ms.prod: visio
api_name:
- Visio.DrawingControl.Window
ms.assetid: 0ecfab32-03eb-e5be-228e-a9e3f96ca536
ms.date: 06/08/2017
---


# DrawingControl.Window Property (Visio)

Returns the window associated with an instance of the Microsoft Visio Drawing Control. Read-only.


## Syntax

 _expression_ . **Window**

 _expression_ A variable that represents a **DrawingControl** object.


### Return Value

Window


## Remarks

For the  **DrawingControl** object, the **Window** property returns the window that the control is displaying. The value is valid only when the control is in place and active.


## Example

The following macro shows how to use the  **Window** property to print the caption of the window associated with the current instance of Visio in the Immediate window.


```vb
 
Public Sub Window_Example() 
 
 Debug.Print Application.Window.Caption 
 
End Sub
```


