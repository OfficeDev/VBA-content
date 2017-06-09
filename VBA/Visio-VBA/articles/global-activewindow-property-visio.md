---
title: Global.ActiveWindow Property (Visio)
keywords: vis_sdr.chm12813035
f1_keywords:
- vis_sdr.chm12813035
ms.prod: visio
api_name:
- Visio.Global.ActiveWindow
ms.assetid: 3d804eb0-9e84-0123-e921-edfbcd79d461
ms.date: 06/08/2017
---


# Global.ActiveWindow Property (Visio)

Returns the active  **Window** object. Read-only.


## Syntax

 _expression_ . **ActiveWindow**

 _expression_ A variable that represents a **Global** object.


### Return Value

Window


## Remarks

The active window can be one of the following window types: Drawing, Stencil, ShapeSheet, Edit Icon, or a Drawing or Stencil window created by an add-on. The application's active window can only be an MDI frame window—it cannot be one of the floating, docked, or anchored windows. For a complete list of window types, see the  **[Type](window-type-property-visio.md)** property.

If a window in an instance of Microsoft Visio is not active, the  **ActiveWindow** property returns **Nothing** .


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to get the active window without qualification from the Microsoft Visio global object, which is automatically available to VBA code that is part of the VBA project of a Visio document.


```vb
 
Public Sub ActiveWindow_Example() 
 
 Dim vsoWindow As Visio.Window 
 
 'Get the active window. 
 Set vsoWindow = ActiveWindow 
 
 'To verify that we got the active window, print its caption. 
 Debug.Print vsoWindow.Caption 
 
End Sub
```


