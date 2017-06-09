---
title: ConnectorFormat.EndConnected Property (Publisher)
keywords: vbapb10.chm3211523
f1_keywords:
- vbapb10.chm3211523
ms.prod: publisher
api_name:
- Publisher.ConnectorFormat.EndConnected
ms.assetid: ace997de-5a11-6b52-ac87-e914adb4212d
ms.date: 06/08/2017
---


# ConnectorFormat.EndConnected Property (Publisher)

Returns an  **MsoTriState** constant indicating whether the end of the specified connector is connected to a shape. Read-only.


## Syntax

 _expression_. **EndConnected**

 _expression_A variable that represents an  **ConnectorFormat** object.


### Return Value

MsoTriState


## Remarks

Use the  **[BeginConnected](connectorformat-beginconnected-property-publisher.md)** property to determine if the beginning of a connector is connected to a shape.

The  **EndConnected** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| The end of the specified connector is not connected to a shape.|
| **msoTriStateMixed**|Return value only; indicates a combination of  **msoTrue** and **msoFalse** in the specified shape range.|
| **msoTrue**| The end of the specified connector is connected to a shape.|

## Example

If the third shape on the first page in the active publication is a connector whose end is connected to a shape, this example stores the connection site number, stores a reference to the connected shape, and then disconnects the end of the connector from the shape.


```vb
Dim intSite As Integer 
Dim shpConnected As Shape 
 
With ActiveDocument.Pages(1).Shapes(3) 
 
 ' Test whether shape is a connector. 
 If .Connector Then 
 With .ConnectorFormat 
 
 ' Test whether connector is connected to another shape. 
 If .End Connected Then 
 
 ' Store connection site number. 
 intSite = .EndConnectionSite 
 
 ' Set reference to connected shape. 
 Set shpConnected = .EndConnectedShape 
 
 ' Disconnect connector and shape. 
 .EndDisconnect 
 End If 
 End With 
 End If 
End With 

```


