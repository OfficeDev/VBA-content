---
title: Shape.AutoShapeType Property (Publisher)
keywords: vbapb10.chm2228274
f1_keywords:
- vbapb10.chm2228274
ms.prod: publisher
api_name:
- Publisher.Shape.AutoShapeType
ms.assetid: f469dc31-a620-5561-ce57-fbff8a5536c0
ms.date: 06/08/2017
---


# Shape.AutoShapeType Property (Publisher)

Returns or sets an  **MsoAutoShapeType**constant that specifies a  **Shape** object's AutoShape type.


## Syntax

 _expression_. **AutoShapeType**

 _expression_A variable that represents a  **Shape** object.


## Remarks

The  **AutoShapeType** property value can be one of the ** [MsoAutoShapeType](http://msdn.microsoft.com/library/7e6fe414-2b25-56d7-a678-b6e718329118%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

AutoShapes correspond to  **Shape** objects, although the **AutoShapeType** property for non-Publisher shapes will also return a value. WordArt, OLE, Web Form control, table and picture frame objects should return **msoShapeMixed** as their **AutoShapeType** property value. Text frames should return **msoShapeRectangle** as their **AutoShapeType** property.


## Example

This example converts the selected  **AutoShape** object to a lightning bolt if it is a heart and to a 5-point star if it is not. For this example to execute properly, you must have an **AutoShape** object selected in the active publication.


```vb
Sub ShapeShift() 
 
 Dim srShift As ShapeRange 
 
 Set srShift = Application.ActiveDocument.Selection.ShapeRange 
 If srShift.AutoShapeType = msoShapeHeart Then 
 srShift.AutoShapeType = msoShapeLightningBolt 
 Else 
 srShift.AutoShapeType = msoShape5pointStar 
 End If 
 
End Sub
```


