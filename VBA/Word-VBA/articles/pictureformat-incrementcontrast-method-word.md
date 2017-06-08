---
title: PictureFormat.IncrementContrast Method (Word)
keywords: vbawd10.chm164298763
f1_keywords:
- vbawd10.chm164298763
ms.prod: word
api_name:
- Word.PictureFormat.IncrementContrast
ms.assetid: afde4afa-53b6-7dd2-57b2-c25a800fb69d
ms.date: 06/08/2017
---


# PictureFormat.IncrementContrast Method (Word)

Changes the contrast of the picture by the specified amount.


## Syntax

 _expression_ . **IncrementContrast**( **_Increment_** )

 _expression_ Required. A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much to change the value of the  **Contrast** property for the picture. A positive value increases the contrast; a negative value decreases the contrast.|

## Remarks

Use the  **Contrast** property to set the absolute contrast for the picture.

You cannot adjust the contrast of a picture past the upper or lower limit for the  **Contrast** property. For example, if the **Contrast** property is initially set to 0.9 and you specify 0.3 for the Increment argument, the resulting contrast level will be 1.0, which is the upper limit for the **Contrast** property, instead of 1.2.


## Example

This example increases the contrast for all embedded OLE objects on the active document that aren't already set to maximum contrast.


```vb
Dim docActive As Document 
Dim shapeLoop As Shape 
 
Set docActive = ActiveDocument 
 
For Each shapeLoop In docActive.Shapes 
 If shapeLoop.Type = msoEmbeddedOLEObject Then 
 shapeLoop.PictureFormat.IncrementContrast 0.1 
 End If 
Next shapeLoop
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

