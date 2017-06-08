---
title: Shape.ScaleHeight Method (PowerPoint)
keywords: vbapp10.chm547010
f1_keywords:
- vbapp10.chm547010
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.ScaleHeight
ms.assetid: 0324449a-535c-e5ec-a9c3-0913f66057c0
ms.date: 06/08/2017
---


# Shape.ScaleHeight Method (PowerPoint)

Scales the height of the shape by a specified factor.


## Syntax

 _expression_. **ScaleHeight**( **_Factor_**, **_RelativeToOriginalSize_**, **_fScale_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Factor_|Required|**Single**|Specifies the ratio between the height of the shape after you resize it and the current or original height. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.|
| _RelativeToOriginalSize_|Required|**MsoTriState**|Specifies whether the shape is scaled relative to its current or original size.|
| _fScale_|Optional|**MsoScaleFrom**|The part of the shape that retains its position when the shape is scaled.|

## Remarks

For pictures and OLE objects, you can indicate whether you want to scale the shape relative to its original size or relative to its current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.

The  _RelativeToOriginalSize_ parameter value can be one of the following **MsoTriState** constants. You can specify **msoTrue** for this parameter only if the specified shape is a picture or an OLE object.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Scales the shape relative to its current size. |
|**msoTrue**|Scales the shape relative to its original size. |
The  _fScale_ parameter value can be one of the following **MsoTriState** constants. The default is **msoScaleFromTopLeft**.


||
|:-----|
|**msoScaleFromBottomRight**|
|**msoScaleFromMiddle**|
|**msoScaleFromTopLeft**|

## Example

This example scales all pictures and OLE objects on  `myDocument` to 175 percent of their original height and width, and it scales all other shapes to 175 percent of their current height and width.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    Select Case s.Type

    Case msoEmbeddedOLEObject, msoLinkedOLEObject, _
            msoOLEControlObject, msoLinkedPicture, msoPicture 

        s.ScaleHeight 1.75, msoTrue
        s.ScaleWidth 1.75, msoTrue

    Case Else

        s.ScaleHeight 1.75, msoFalse
        s.ScaleWidth 1.75, msoFalse

    End Select

Next
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

