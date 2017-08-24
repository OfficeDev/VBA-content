---
title: Shapes.AddEmptyPictureFrame Method (Publisher)
keywords: vbapb10.chm2162757
f1_keywords:
- vbapb10.chm2162757
ms.prod: publisher
api_name:
- Publisher.Shapes.AddEmptyPictureFrame
ms.assetid: e473dea8-6d94-e9e4-ddb6-27c1fc8930e8
ms.date: 06/08/2017
---


# Shapes.AddEmptyPictureFrame Method (Publisher)

Returns a  **Shape** object that represents an empty picture frame inserted at the specified coordinates.


## Syntax

 _expression_. **AddEmptyPictureFrame**( **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Left|Required| **Variant**|The position of the left edge of the shape representing the picture.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the picture.|
|Width|Optional| **Variant**|The width of the shape representing the picture. Default is -1, meaning that the width of the shape is automatically set to 72 points if the parameter is left blank.|
|Height|Optional| **Variant**|The height of the shape representing the picture. Default is -1, meaning that the height of the shape is automatically set to 54 points if the parameter is left blank.|

### Return Value

Shape


## Remarks

For  **Left**,  **Top**,  **Width**, and  **Height** arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "1.5 in").

The blank picture frame has the default ToolTip "Empty Picture Frame". This is changed to "Picture" when an image is selected for the  **Shape**.


## Example

This example places an empty picture frame in the center of the first page of the publication and rotates it by 45 degrees. The  **AlternativeText** property is set to "Picture Placeholder 1" for the Web.


```vb
Dim shpPlaceholder As Shape 
 
Set shpPlaceholder = _ 
 ActiveDocument.Pages(1).Shapes.AddEmptyPictureFrame( _ 
 230, 320, 150, 150) 
 
With shpPlaceholder 
 .AlternativeText = "Picture Placeholder 1" 
 .Rotation = 45 
End With 
 

```


