---
title: Application.ObjectType Property (Visio)
keywords: vis_sdr.chm10013960
f1_keywords:
- vis_sdr.chm10013960
ms.prod: visio
api_name:
- Visio.Application.ObjectType
ms.assetid: 50baf864-034e-9051-3671-a3c8f0f368ed
ms.date: 06/08/2017
---


# Application.ObjectType Property (Visio)

Returns an object's type. Read-only.


## Syntax

 _expression_ . **ObjectType**

 _expression_ A variable that represents an **Application** object.


### Return Value

Integer


## Remarks

Constants representing object types are prefixed with  **visObjType** and are declared by the Visio type library in **[VisObjectTypes](visobjecttypes-enumeration-visio.md)** .


## Example

This example shows how to use the  **ObjectType** property of a page to iterate recursively through a group and identify the top shape.


```vb
Public Sub ObjectType_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoPage As Visio.Page 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Pentagon"), 3#, 8.5 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Ellipse"), 3#, 7.625 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Rounded rectangle"), 3#, 7# 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Circle"), 3#, 6.25 
 
 Application.ActiveWindow.SelectAll 
 
 ActiveWindow.DeselectAll 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Circle"), visSelect 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Rounded rectangle"), visSelect 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Ellipse"), visSelect 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Pentagon"), visSelect 
 ActiveWindow.Selection.Group 
 
 Set vsoPage = ActivePage 
 Set vsoShapes = vsoPage.Shapes 
 Set vsoShape = vsoShapes.Item(2) 
 
 Call GetTopShape(vsoShape) 
 
End Sub 
 
Function GetTopShape(vsoShape As Visio.Shape) As String 
 
 Dim vsoShapeParent As Object 
 Dim vsoShapeParentParent As Object 
 
 Set vsoShapeParent = vsoShape.Parent 
 
 If vsoShapeParent.ObjectType = visObjTypeShape Then 
 
 Set vsoShapeParentParent = vsoShapeParent.Parent 
 
 'If vsoShapeParent's parent isn't a page, keep going up. 
 If vsoShapeParentParent.ObjectType = visObjTypePage Then 
 GetTopShape = vsoShapeParent.Name 
 Else 
 GetTopShape = GetTopShape(vsoShapeParent) 
 End If 
 
 End If 
 
 Debug.Print vsoShapeParent.Name 
 
End Function
```


