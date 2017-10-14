---
title: Shapes.AddConnector Method (Project)
ms.prod: project-server
ms.assetid: bfd75cf3-f70b-8d19-bf28-94e2f4b227dd
ms.date: 06/08/2017
---


# Shapes.AddConnector Method (Project)
Creates a connector and returns a  **Shape** object the represents the new connector.

## Syntax

 _expression_. **AddConnector** _(Type,_ _BeginX,_ _BeginY,_ _EndX,_ _EndY)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoConnectorType**|The type of connector. Can be one of the following constants:  **msoConnectorElbow**,  **msoConnectorTypeMixed**,  **msoConnectorCurve**, or  **msoConnectorStraight**.|
| _BeginX_|Required|**Single**|The horizontal position (in points) of the connector's starting point, relative to the upper-left corner of the document.|
| _BeginY_|Required|**Single**|The vertical position (in points) of the connector's starting point.|
| _EndX_|Required|**Single**|The horizontal position (in points) of the connector's end point.|
| _EndY_|Required|**Single**|The vertical position (in points) of the connector's end point.|
| _Type_|Required|MSOCONNECTORTYPE||
| _BeginX_|Required|FLOAT||
| _BeginY_|Required|FLOAT||
| _EndX_|Required|FLOAT||
| _EndY_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Shape**


## Remarks


 **Note**  In Project, the methods to attach the beginning and end of a connector to other shapes in the report ( **ConnectorFormat.BeginConnect** and **ConnectorFormat.EndConnect**) do not work. You can use only the  **AddConnector** parameters to position the connector. For more information, see the[ConnectorFormat](shape-connectorformat-property-project.md) property.


## Example

The following example creates a report that contains two cloud shapes, and then adds a blue-green curved connector line that is two points wide.


```vb
Sub ConnectClouds()
    Dim shapeReport As Report
    Dim reportName As String
    Dim connectorShape As shape
    
    ' Add a report.
    reportName = "Cloud report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)

    ' Add two clouds.
    Dim cloudShape1 As shape
    Dim cloudShape2 As shape
    Set cloudShape1 = shapeReport.Shapes.AddShape(msoShapeCloud, 20, 20, 100, 60)
    Set cloudShape2 = shapeReport.Shapes.AddShape(msoShapeCloud, 100, 200, 60, 100)
    
    Set connectorShape = shapeReport.Shapes.AddConnector(msoConnectorCurve, 80, 80, 130, 200)
        
    With connectorShape
        .Line.Weight = 2
        .Line.ForeColor.RGB = &;HAAFF00
    End With
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
[ConnectorFormat Property](shape-connectorformat-property-project.md)
[AutoShapeType Property](shape-autoshapetype-property-project.md)
[MsoConnectorType](http://msdn.microsoft.com/en-us/library/office/ff860918%28v=office.15%29)
