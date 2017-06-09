---
title: ConnectorFormat.BeginConnected Property (Publisher)
keywords: vbapb10.chm3211520
f1_keywords:
- vbapb10.chm3211520
ms.prod: publisher
api_name:
- Publisher.ConnectorFormat.BeginConnected
ms.assetid: ed70561e-b63e-530d-87be-1e6b7d87c425
ms.date: 06/08/2017
---


# ConnectorFormat.BeginConnected Property (Publisher)

Returns an  **MsoTriState**constant indicating whether the beginning of the specified connector is connected to a shape. Read-only.


## Syntax

 _expression_. **BeginConnected**

 _expression_A variable that represents a  **ConnectorFormat** object.


### Return Value

MsoTriState


## Remarks

The  **BeginConnected** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

Use the  **[EndConnected](connectorformat-endconnected-property-publisher.md)** property to determine if the end of a connector is connected to a shape.


## Example

If the third shape on the first page in the active publication is a connector whose beginning is connected to a shape, this example stores the connection site number, stores a reference to the connected shape, and then disconnects the beginning of the connector from the shape.


```vb
Dim intSite As Integer 
Dim shpConnected As Shape 
 
With ActiveDocument.Pages(1).Shapes(3) 
 
 ' Test whether shape is a connector. 
 If .Connector Then 
 With .ConnectorFormat 
 
 ' Test whether connector is connected to another shape. 
 If .BeginConnected Then 
 
 ' Store connection site number. 
 intSite = .BeginConnectionSite 
 
 ' Set reference to connected shape. 
 Set shpConnected = .BeginConnectedShape 
 
 ' Disconnect connector and shape. 
 .BeginDisconnect 
 End If 
 End With 
 End If 
End With 

```


