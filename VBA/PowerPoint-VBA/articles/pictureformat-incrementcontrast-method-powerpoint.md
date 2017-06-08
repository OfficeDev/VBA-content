---
title: PictureFormat.IncrementContrast Method (PowerPoint)
keywords: vbapp10.chm551003
f1_keywords:
- vbapp10.chm551003
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.IncrementContrast
ms.assetid: ad5c45b2-0193-eda9-a511-4dd9050daee7
ms.date: 06/08/2017
---


# PictureFormat.IncrementContrast Method (PowerPoint)

Changes the contrast of the picture by the specified amount. 


## Syntax

 _expression_. **IncrementContrast**( **_Increment_** )

 _expression_ A variable that represents an **PictureFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how much to change the value of the  **Contrast** property for the picture. A positive value increases the contrast; a negative value decreases the contrast.|

## Remarks

Use the  **[Contrast](pictureformat-contrast-property-powerpoint.md)** property to set the absolute contrast for the picture.

You cannot adjust the contrast of a picture past the upper or lower limit for the  **Contrast** property. For example, if the **Contrast** property is initially set to 0.9 and you specify 0.3 for the Increment argument, the resulting contrast level will be 1.0, which is the upper limit for the **Contrast** property, instead of 1.2.


## Example

This example increases the contrast for all pictures on  `myDocument` that aren't already set to maximum contrast.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.Type = msoPicture Then

        s.PictureFormat.IncrementContrast 0.1

    End If

Next
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-powerpoint.md)

