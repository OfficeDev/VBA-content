---
title: Shapes.AddConnector Method (Publisher)
keywords: vbapb10.chm2162705
f1_keywords:
- vbapb10.chm2162705
ms.prod: publisher
api_name:
- Publisher.Shapes.AddConnector
ms.assetid: fd1ef969-7960-2555-e355-9804c86f6c01
ms.date: 06/08/2017
---


# Shapes.AddConnector Method (Publisher)

Adds a new  **[Shape](shape-object-publisher.md)** object representing a connector to the specified **[Shapes](shapes-object-publisher.md)** collection.


## Syntax

 _expression_. **AddConnector**( **_Type_**,  **_BeginX_**,  **_BeginY_**,  **_EndX_**,  **_EndY_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Required| **MsoConnectorType**|The type of connector to add.|
|BeginX|Required| **Variant**|The x-coordinate of the beginning point of the connector.|
|BeginY|Required| **Variant**|The y-coordinate of the beginning point of the connector.|
|EndX|Required| **Variant**|The x-coordinate of the ending point of the connector.|
|EndY|Required| **Variant**|The y-coordinate of the ending point of the connector.|

### Return Value

Shape


## Remarks

For the BeginX, BeginY, EndX, and EndY parameters, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

The new connector isn't connected to any other shape; use the  **[BeginConnect](connectorformat-beginconnect-method-publisher.md)** and  **[EndConnect](connectorformat-endconnect-method-publisher.md)** methods to connect the new connector to another shape.

The Type parameter can be one of these  **MsoConnectorType** constants.



| **msoConnectorCurve**|Adds a curved connector.|
| **msoConnectorElbow**|Adds an elbow-shaped connector.|
| **msoConnectorStraight**|Adds a straight-line connector.|
| **msoConnectorTypeMixed**|Not used with this method.|

## Example

The following example adds a new straight-line connector to the first page of the active publication.


```vb
Dim shpConnect As Shape 
 
Set shpConnect = ActiveDocument.Pages(1).Shapes.AddConnector _ 
 (Type:=msoConnectorStraight, _ 
 BeginX:=144, BeginY:=144, _ 
 EndX:=180, EndY:=72)
```


