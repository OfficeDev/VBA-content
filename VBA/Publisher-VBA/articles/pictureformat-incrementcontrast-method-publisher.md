---
title: PictureFormat.IncrementContrast Method (Publisher)
keywords: vbapb10.chm3604497
f1_keywords:
- vbapb10.chm3604497
ms.prod: publisher
api_name:
- Publisher.PictureFormat.IncrementContrast
ms.assetid: cff50058-2b88-fc2d-633d-411380e5f2f3
ms.date: 06/08/2017
---


# PictureFormat.IncrementContrast Method (Publisher)

Changes the contrast of the picture by the specified amount.


## Syntax

 _expression_. **IncrementContrast**( **_Increment_**)

 _expression_A variable that represents a  **PictureFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Single**|Specifies how much to change the value of the  **[Contrast](pictureformat-contrast-property-publisher.md)** property for the picture. A positive value increases the contrast; a negative value decreases the contrast. Valid values are between - 1 and 1.|

## Remarks

You cannot adjust the contrast of a picture past the upper or lower limit for the  **Contrast** property. For example, if the **Contrast** property is initially set to 0.9 and you specify 0.3 for the **_Increment_** argument, the resulting contrast level will be 1.0, which is the upper limit for the **Contrast** property, instead of 1.2.

Use the  **Contrast** property to set the absolute contrast for the picture.


## Example

This example increases the contrast for all pictures on the first page of the active publication that aren't already set to maximum contrast.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = msoPicture Then 
 shpLoop.PictureFormat.IncrementContrast Increment:=0.1 
 End If 
Next shpLoop 

```


