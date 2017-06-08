---
title: Shapes.AddShape Method (Publisher)
keywords: vbapb10.chm2162712
f1_keywords:
- vbapb10.chm2162712
ms.prod: publisher
api_name:
- Publisher.Shapes.AddShape
ms.assetid: 500d8cb3-f066-fdb6-09ac-b03c7822e8bd
ms.date: 06/08/2017
---


# Shapes.AddShape Method (Publisher)

Adds a new  **Shape** object representing an AutoShape to the specified **Shapes** collection.


## Syntax

 _expression_. **AddShape**( **_Type_**,  **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Required| **MsoAutoShapeType**|The type of AutoShape to draw. For a complete list of MsoAutoShapeType constants, see the Object Browser.|
|Left|Required| **Variant**|The position of the left edge of the shape representing the AutoShape.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the AutoShape.|
|Width|Required| **Variant**|The width of the shape representing the AutoShape.|
|Height|Required| **Variant**|The height of the shape representing the AutoShape.|

### Return Value

Shape


## Remarks

For the  **_Left_**,  **_Top_**,  **_Width_**, and  **_Height_** arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

The following example adds a rectangle to the first page of the active publication.


```vb
Dim shpShape As Shape 
 
Set shpShape = ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeRectangle, _ 
 Left:=144, Top:=144, _ 
 Width:=72, Height:=144) 

```


