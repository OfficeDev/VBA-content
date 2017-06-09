---
title: Shape.ConnectorFormat Property (Project)
ms.prod: project-server
ms.assetid: 8bcbe86a-164e-038f-c41a-2d951e549aef
ms.date: 06/08/2017
---


# Shape.ConnectorFormat Property (Project)
Gets a  **ConnectorFormat** object that contains connector formatting properties. Applies to a **Shape** that represents a connector. Read-only **[ConnectorFormat](http://msdn.microsoft.com/en-us/library/office/ff820940%28v=office.15%29)**.

## Syntax

 _expression_. **ConnectorFormat**

 _expression_ A variable that represents a **Shape** object.


## Remarks


 **Note**  In Project, the connect and disconnect methods do not work for a  **ConnectorFormat** object. So, the **RerouteConnections** method and the **BeginConnected**,  **BeginConnectedShape**,  **BeginConnectedSite**,  **EndConnected**,  **EndConnectedShape**, and  **EndConnectedSite** properties have no meaning.

For example, in the following code snippet, the  **BeginConnect** method gives a run-time error 13, 'Type mismatch'.


```vb
Set connectorShape = oReport.Shapes.AddConnector(msoConnectorCurve, 100, 250, 150, 280)

With connectorShape
    ' Type mismatch error:
    .ConnectorFormat.BeginConnect ConnectedShape:=oReport.Shapes(5), _
        ConnectionSite:=1
    .ConnectorFormat.EndConnect ConnectedShape:=oReport.Shapes(6),_
        ConnectionSite:=1
End With
```


## Property value

 **CONNECTORFORMAT**


## See also


#### Other resources


[Shape Object](shape-object-project.md)
[AddConnector Method](shapes-addconnector-method-project.md)
[ConnectorFormat](http://msdn.microsoft.com/en-us/library/office/ff820940%28v=office.15%29)
