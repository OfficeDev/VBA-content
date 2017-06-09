---
title: Shape.ScaleHeight Method (Word)
keywords: vbawd10.chm161480723
f1_keywords:
- vbawd10.chm161480723
ms.prod: word
api_name:
- Word.Shape.ScaleHeight
ms.assetid: 994aac8b-5842-5986-0d27-01e52e01066d
ms.date: 06/08/2017
---


# Shape.ScaleHeight Method (Word)

Scales the height of the shape by a specified factor.


## Syntax

 _expression_ . **ScaleHeight**( **_Factor_** , **_RelativeToOriginalSize_** , **_Scale_** )

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Factor_|Required| **Single**|Specifies the ratio between the height of the shape after you resize it and the current or original height. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.|
| _RelativeToOriginalSize_|Required| **MsoTriState**| **True** to scale the shape relative to its original size. **False** to scale it relative to its current size. You can specify **True** for this argument only if the specified shape is a picture or an OLE object.|
| _Scale_|Optional| **MsoScaleFrom**|The part of the shape that retains its position when the shape is scaled.|

## Remarks

For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original size or relative to the current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.


## Example

This example scales all pictures and OLE objects on  _myDocument_ to 175 percent of their original height and width, and it scales all other shapes to 175 percent of their current height and width.


```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 Select Case s.Type 
 Case msoEmbeddedOLEObject, msoLinkedOLEObject, _ 
 msoOLEControlObject, _ 
 msoLinkedPicture, msoPicture 
 s.ScaleHeight 1.75, True 
 s.ScaleWidth 1.75, True 
 Case Else 
 s.ScaleHeight 1.75, False 
 s.ScaleWidth 1.75, False 
 End Select 
Next
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

