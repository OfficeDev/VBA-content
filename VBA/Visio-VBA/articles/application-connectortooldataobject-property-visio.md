---
title: Application.ConnectorToolDataObject Property (Visio)
keywords: vis_sdr.chm10051835
f1_keywords:
- vis_sdr.chm10051835
ms.prod: visio
api_name:
- Visio.Application.ConnectorToolDataObject
ms.assetid: 7b1eedad-3d62-c2a1-5ba7-200a594ba32f
ms.date: 06/08/2017
---


# Application.ConnectorToolDataObject Property (Visio)

Returns an  **IDataObject** interface representing the active **Connector** tool used in the Microsoft Visio user interface. Read-only.


## Syntax

 _expression_ . **ConnectorToolDataObject**

 _expression_ A variable that represents an **Application** object.


### Return Value

IDataObject


## Remarks

 By default, **ConnectorToolDataObject** returns the built-in Visio **Connector** tool. If a master from a stencil is the active connector, **ConnectorToolDataObject** returns a data object for that master. If Visio fails to retrieve the internal **IDataObject** , it raises an exception.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ConnectorToolDataObject** property to connect two shapes. It drops two masters on the page and connects them with a **Dynamic Connector** shape, using dynamic glue. Before running this macro, open the **Basic Shapes** stencil, if it is not already open.


```vb
Public Sub ConnectorToolDataObject_Example() 
 
 Dim vsoSquareShape As Visio.Shape 
 Dim vsoCircleShape As Visio.Shape 
 Dim vsoConnectorShape As Visio.Shape 
 
 Dim vsoCell1 As Visio.Cell 
 Dim vsoCell2 As Visio.Cell 
 Dim vsoCell3 As Visio.Cell 
 Dim vsoCell4 As Visio.Cell 
 
 Set vsoSquareShape = ActiveWindow.Page.Drop(Documents("BASIC_U.VSS").Masters.ItemU("Square"), 4, 9) 
 Set vsoCircleShape = ActiveWindow.Page.Drop(Documents("BASIC_U.VSS").Masters.ItemU("Circle"), 4#, 6) 
 Set vsoConnectorShape = Application.ActiveWindow.Page.Drop(Application.ConnectorToolDataObject, 2, 2) 
 
 Set vsoCell1 = ActivePage.Shapes(3).Cells("BeginX") 
 Set vsoCell2 = ActivePage.Shapes(1).CellsSRC(7, 0, 0) 
 vsoCell1.GlueTo vsoCell2 
 
 Set vsoCell3 = ActivePage.Shapes(3).Cells("EndX") 
 Set vsoCell4 = ActivePage.Shapes(2).CellsSRC(7, 1, 0) 
 vsoCell3.GlueTo vsoCell4 
 
 
End Sub
```


