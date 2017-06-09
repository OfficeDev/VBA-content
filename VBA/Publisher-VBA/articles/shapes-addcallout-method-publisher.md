---
title: Shapes.AddCallout Method (Publisher)
keywords: vbapb10.chm2162704
f1_keywords:
- vbapb10.chm2162704
ms.prod: publisher
api_name:
- Publisher.Shapes.AddCallout
ms.assetid: bbf5f913-fcf0-b700-0c7e-9f0bdc7c6aea
ms.date: 06/08/2017
---


# Shapes.AddCallout Method (Publisher)

Adds a new  **[Shape](shape-object-publisher.md)** object representing a borderless line callout to the specified **[Shapes](shapes-object-publisher.md)** collection.


## Syntax

 _expression_. **AddCallout**( **_Type_**,  **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Required| **MsoCalloutType**|The type of callout line.|
|Left|Required| **Variant**|The position of the left edge of the shape representing the line callout.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the line callout.|
|Width|Required| **Variant**|The width of the shape representing the line callout.|
|Height|Required| **Variant**|The height of the shape representing the line callout.|

### Return Value

Shape


## Remarks

For the Left, Top, Width, and Height arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

The Type parameter can be one of these  **MsoCalloutType** constants.



| **msoCalloutOne**|A horizontal or vertical single-segment callout line.|
| **msoCalloutTwo**|A freely-rotating single-segment callout line.|
| **msoCalloutThree**|A two-segment callout line.|
| **msoCalloutFour**|A three-segment callout line.|

## Example

The following example adds a new freely-rotating callout line to the first page of the active publication.


```vb
Dim shpCallout As Shape 
 
Set shpCallout = ActiveDocument.Pages(1).Shapes.AddCallout _ 
 (Type:=msoCalloutTwo, _ 
 Left:=144, Top:=216, _ 
 Width:=36, Height:=72)
```


