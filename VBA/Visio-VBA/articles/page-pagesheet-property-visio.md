---
title: Page.PageSheet Property (Visio)
keywords: vis_sdr.chm10914000
f1_keywords:
- vis_sdr.chm10914000
ms.prod: visio
api_name:
- Visio.Page.PageSheet
ms.assetid: 495709a8-92f0-6fdf-753f-7ac25c5daaab
ms.date: 06/08/2017
---


# Page.PageSheet Property (Visio)

Returns the page sheet (an object that represents the ShapeSheet spreadsheet) of a page. Read-only.


## Syntax

 _expression_ . **PageSheet**

 _expression_ A variable that represents a **Page** object.


### Return Value

Shape


## Remarks

Every page contains a tree of  **Shape** objects. Constants representing shape types are prefixed with **visType** and are declared by the Visio type library in **[VisShapeTypes](visshapetypes-enumeration-visio.md)** .

In the tree of shapes of a page, there is exactly one shape of type  **visTypePage** . This shape is always the root shape in the tree, and the **PageSheet** property returns this shape.

The page sheet contains important settings for the page such as its size and scale. It also contains the Layers section that defines the layers for that page.

An alternative way to obtain a page's page shape is to use the following macro:




```vb
Sub PagePageSheet_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoPage As Visio.Page 
 Set vsoPage = ActivePage 
 Set vsoShapes = vsoPage.Shapes 
 Set vsoShape = vsoShapes("ThePage") 
 
End Sub
```


