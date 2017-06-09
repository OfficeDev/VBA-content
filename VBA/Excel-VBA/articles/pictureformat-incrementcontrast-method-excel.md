---
title: PictureFormat.IncrementContrast Method (Excel)
keywords: vbaxl10.chm113021
f1_keywords:
- vbaxl10.chm113021
ms.prod: excel
api_name:
- Excel.PictureFormat.IncrementContrast
ms.assetid: 6bb72eed-c291-fac2-f4ca-4ca847bd8458
ms.date: 06/08/2017
---


# PictureFormat.IncrementContrast Method (Excel)

Changes the contrast of the picture by the specified amount. Use the  **[Contrast](pictureformat-contrast-property-excel.md)** property to set the absolute contrast for the picture.


## Syntax

 _expression_ . **IncrementContrast**( **_Increment_** )

 _expression_ A variable that represents a **PictureFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much to change the value of the  **Contrast** property for the picture. A positive value increases the contrast; a negative value decreases the contrast.|

## Remarks

You cannot adjust the contrast of a picture past the upper or lower limit for the  **Contrast** property. For example, if the **Contrast** property is initially set to 0.9 and you specify 0.3 for the _Increment_ argument, the resulting contrast level will be 1.0, which is the upper limit for the **Contrast** property, instead of 1.2.


## Example

This example increases the contrast for all pictures on  `myDocument` that aren?t already set to maximum contrast.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 If s.Type = msoPicture Or s.Type = msoLinkedPicture Then 
 s.PictureFormat.IncrementContrast 0.1 
 End If 
Next
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-excel.md)

