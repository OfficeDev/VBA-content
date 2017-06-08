---
title: Window.Type Property (Visio)
keywords: vis_sdr.chm11614595
f1_keywords:
- vis_sdr.chm11614595
ms.prod: visio
api_name:
- Visio.Window.Type
ms.assetid: 92dd1e1e-2acc-d918-aab6-f267ecc18c26
ms.date: 06/08/2017
---


# Window.Type Property (Visio)

Returns the type of the object. Read-only.


## Syntax

 _expression_ . **Type**

 _expression_ A variable that represents a **Window** object.


### Return Value

Integer


## Remarks

Type value constants for  **Window** objects (the possible values that the **Type** property of a **Window** object returns) are declared by the Visio type library in **[VisWinTypes](viswintypes-enumeration-visio.md)** .

If a  **Window** object is type **visDrawing** , use the **SubType** property to determine the type of drawing window represented by the object.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Type** property to determine the type of a window.


```vb
 
Public Sub Type_Example() 
 
 Dim vsoMaster As Visio.Master 
 Dim vsoShape As Visio.Shape 
 Dim vsoIconWindow As Visio.Window 
 Dim vsoShapeSheetWindow As Visio.Window 
 Dim vsoStencilWindow As Visio.Window 
 
 'Draw a shape. 
 Set vsoShape = ActivePage.DrawRectangle(1, 1, 2, 3) 
 
 'Open the document stencil window. 
 Set vsoStencilWindow = ThisDocument.OpenStencilWindow 
 
 'Open the ShapeSheet window of vsoShape. 
 Set vsoShapeSheetWindow = vsoShape.OpenSheetWindow 
 
 'Add a master to the document stencil and open its icon editing window. 
 Set vsoMaster = ThisDocument.Masters.Add 
 Set vsoIconWindow = vsoMaster.OpenIconWindow 
 
 'Use the Type property to verify each window's type. 
 'This will print 7, 3, and 4 in the Immediate window to indicate 
 'a docked, built-in stencil window; a ShapeSheet window; and an 
 'icon editing window, respectively. 
 Debug.Print vsoStencilWindow.Type 
 Debug.Print vsoShapeSheetWindow.Type 
 Debug.Print vsoIconWindow.Type 
 
End Sub
```


