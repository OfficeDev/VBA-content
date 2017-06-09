---
title: Master.PageSheet Property (Visio)
keywords: vis_sdr.chm10714000
f1_keywords:
- vis_sdr.chm10714000
ms.prod: visio
api_name:
- Visio.Master.PageSheet
ms.assetid: 8ec4d38a-79fe-018d-9bc8-3a9c0221f018
ms.date: 06/08/2017
---


# Master.PageSheet Property (Visio)

Returns the page sheet (an object that represents the ShapeSheet spreadsheet) of a master. Read-only.


## Syntax

 _expression_ . **PageSheet**

 _expression_ A variable that represents a **Master** object.


### Return Value

Shape


## Remarks

Every master contains a tree of  **Shape** objects. Constants representing shape types are prefixed with **visType** and are declared by the Visio type library in **[VisShapeTypes](visshapetypes-enumeration-visio.md)** .

In the tree of shapes of a master, there is exactly one shape of type  **visTypePage** . This shape is always the root shape in the tree, and the **PageSheet** property returns this shape.

The page sheet contains important settings for the master such as its size and scale. It also contains the Layers section that defines the layers for that master.

Assuming that there is at least one shape on the page, you can use the following macro to get a master's page shape:




```vb
Sub MasterPageSheet_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoMaster As Visio.Master 
 Set vsoMaster = ActiveDocument.Masters.Item(1) 
 Set vsoShapes = vsoMaster.Shapes 
 Set vsoShape = vsoShapes("ThePage") 
 
End Sub
```


