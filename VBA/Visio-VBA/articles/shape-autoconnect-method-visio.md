---
title: Shape.AutoConnect Method (Visio)
keywords: vis_sdr.chm11260240
f1_keywords:
- vis_sdr.chm11260240
ms.prod: visio
api_name:
- Visio.Shape.AutoConnect
ms.assetid: 36b634be-9943-1aec-f8e0-70467b82eed1
ms.date: 06/08/2017
---


# Shape.AutoConnect Method (Visio)

Automatically draws a connection in the specified direction between the shape and another shape on the drawing page.


## Syntax

 _expression_ . **AutoConnect**( **_ToShape_** , **_PlacementDir_** , **_Connector_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ToShape_|Required| **[IVSHAPE]**|The shape to draw the connection to.|
| _PlacementDir_|Required| **VisAutoConnectDir**|The direction in which to draw the connection. See Remarks for possible values.|
| _Connector_|Optional| **[UNKNOWN]**|The connector to use.|

### Return Value

Nothing


## Remarks

The  **AutoConnect** method lets you automatically draw connections between shapes on the drawing page while specifying the direction of the connection and, optionally, the connector.

For the ToShape parameter, pass the  **Shape** object to which you want to draw the connection.

For the PlacementDir parameter, pass a value from the  **VisAutoConnectDir** enumeration to specify the connection direction?that is, where to locate the connected shape with respect to the primary shape. Possible values for PlacementDir are as follows.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|visAutoConnectDirDown|2|Connect down.|
|visAutoConnectDirLeft|3|Connect to the left.|
|visAutoConnectDirNone|0|Connect without relocating the shapes.|
|visAutoConnectDirRight|4|Connect to the right|
|visAutoConnectDirUp|1|Connect up.|
If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVShape.AutoConnect(Microsoft.Office.Interop.Visio.Shape, Microsoft.Office.Interop.Visio.VisAutoConnectDir, object)**
    

## Example

 The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **AutoConnect** method to draw a connection between two flowchart shapes, a decision shape and a process shape, by using a third shape, a dynamic connector, all of which were added to an empty drawing page from the Basic Flowchart Shapes (US Units) stencil.

Because the example calls the method on the decision shape, Visio draws the connector from the decision shape to the process shape. Because we pass the method the enumerated value  **visAutoConnectDirRight** for the PlacementDir parameter, Visio places the process shape automatically to the right of the decision shape on the drawing page, regardless of its previous location.




```vb
Public Sub AutoConnect_Example() 
 
    Dim vsoShape1 As Visio.Shape 
    Dim vsoShape2 As Visio.Shape 
    Dim vsoConnectorShape As Visio.Shape 
 
    Set vsoShape1 = Visio.ActivePage.Shapes("Decision") 
    Set vsoShape2 = Visio.ActivePage.Shapes("Process") 
    Set vsoConnectorShape = Visio.ActivePage.Shapes("Dynamic connector") 
 
    vsoShape1.AutoConnect vsoShape2, visAutoConnectDirRight, vsoConnectorShape 
 
End Sub
```


