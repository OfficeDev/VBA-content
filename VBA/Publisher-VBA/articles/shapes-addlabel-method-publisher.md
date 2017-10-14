---
title: Shapes.AddLabel Method (Publisher)
keywords: vbapb10.chm2162707
f1_keywords:
- vbapb10.chm2162707
ms.prod: publisher
api_name:
- Publisher.Shapes.AddLabel
ms.assetid: 5a803aa2-d37f-6da1-7d8b-58ee2dcd8146
ms.date: 06/08/2017
---


# Shapes.AddLabel Method (Publisher)

Adds a new  **[Shape](shape-object-publisher.md)** object representing a text label to the specified **[Shapes](shapes-object-publisher.md)** collection.


## Syntax

 _expression_. **AddLabel**( **_Orientation_**,  **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Orientation|Required| **PbTextOrientation**|The orientation of the label.|
|Left|Required| **Variant**|The position of the left edge of the shape representing the text label.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the text label.|
|Width|Required| **Variant**|The width of the shape representing the text label.|
|Height|Required| **Variant**|The height of the shape representing the text label.|

### Return Value

Shape


## Remarks

For the Left, Top, Width, and Height arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

The Orientation parameter can be one of these  **PbTextOrientation** constants.



| **pbTextOrientationHorizontal**|A horizontal text label for left-to-right languages.|
| **pbTextOrientationRightToLeft**| A horizontal text label for right-to-left languages.|
| **pbTextOrientationVerticalEastAsia**|A vertical text label for East Asian languages.|

## Example

The following example adds a new horizontal text label to the first page of the active publication.


```vb
Dim shpLabel As Shape 
 
Set shpLabel = ActiveDocument.Pages(1).Shapes.AddLabel _ 
 (Orientation:=pbTextOrientationHorizontal, _ 
 Left:=144, Top:=144, _ 
 Width:=72, Height:=18)
```


