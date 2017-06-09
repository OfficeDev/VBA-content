---
title: Shapes.AddLine Method (Publisher)
keywords: vbapb10.chm2162708
f1_keywords:
- vbapb10.chm2162708
ms.prod: publisher
api_name:
- Publisher.Shapes.AddLine
ms.assetid: 43df8878-5640-875f-06e0-37e1feb47b78
ms.date: 06/08/2017
---


# Shapes.AddLine Method (Publisher)

Adds a new  **[Shape](shape-object-publisher.md)** object representing a line to the specified **[Shapes](shapes-object-publisher.md)** collection.


## Syntax

 _expression_. **AddLine**( **_BeginX_**,  **_BeginY_**,  **_EndX_**,  **_EndY_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|BeginX|Required| **Variant**|The x-coordinate of the beginning point of the line.|
|BeginY|Required| **Variant**|The y-coordinate of the beginning point of the line.|
|EndX|Required| **Variant**|The x-coordinate of the ending point of the line.|
|EndY|Required| **Variant**|The y-coordinate of the ending point of the line.|

### Return Value

Shape


## Remarks

For the  **_BeginX_**,  **_BeginY_**,  **_EndX_**, and  **_EndY_** arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

The following example adds a new line to the first page of the active publication.


```vb
Dim shpLine As Shape 
 
Set shpLine = ActiveDocument.Pages(1).Shapes.AddLine _ 
 (BeginX:=144, BeginY:=144, _ 
 EndX:=180, EndY:=72) 

```


