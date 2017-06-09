---
title: Shape.Type Property (Publisher)
keywords: vbapb10.chm2228307
f1_keywords:
- vbapb10.chm2228307
ms.prod: publisher
api_name:
- Publisher.Shape.Type
ms.assetid: bb712dd4-5d81-10e0-9b4c-4af6a09a3c71
ms.date: 06/08/2017
---


# Shape.Type Property (Publisher)

Specifies the shape type. Read-only.


## Syntax

 _expression_. **Type**

 _expression_A variable that represents a  **Shape** object.


## Remarks

The  **Type** property value can be one of the **[PbShapeType](pbshapetype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example formats the callout type of the specified shape if the shape is a callout. This example assumes there is at least one shape on the first page of the active publication.


```vb
Sub SetCalloutType() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = pbCallout Then 
 With .Callout 
 .Border = msoTrue 
 .Type = msoCalloutThree 
 End With 
 End If 
 End With 
End Sub
```


