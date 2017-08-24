---
title: ShapeRange.SaveAsPicture Method (Publisher)
keywords: vbapb10.chm2294050
f1_keywords:
- vbapb10.chm2294050
ms.prod: publisher
api_name:
- Publisher.ShapeRange.SaveAsPicture
ms.assetid: 0be9b741-8f11-a386-313b-231a3269883a
ms.date: 06/08/2017
---


# ShapeRange.SaveAsPicture Method (Publisher)

Saves a range of one or more shapes as a picture file.


## Syntax

 _expression_. **SaveAsPicture**( **_Filename_**,  **_pbResolution_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Filename|Required| **String**|The path and file name of the new picture file you want to create. The graphics format the picture is saved in is determined by the file name extension (such as .jpg or .gif) you specify.|
|pbResolution|Optional| **PbPictureResolution**|The resolution in which you want the picture to be saved. See Remarks for possible values.|

## Remarks

Possible values for the pbResolution parameter are declared in the  **[PbPictureResolution](pbpictureresolution-enumeration-publisher.md)** enumeration in the Microsoft Publisher type library.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SaveAsPicture** method to save all the shapes on the first page of the active publication as a .jpg picture file.

Before running this code, replace  _filename.jpg_ with a valid file name and the path to a folder on your computer where you have permission to save files.




```vb
Public Sub SaveAsPicture_Example() 
 
 Dim pubShapeRange As Publisher.ShapeRange 
 Set pubShapeRange = ThisDocument.Pages(1).Shapes.Range 
 
 pubShapeRange.SaveAsPicture "filename.jpg" 
 
End Sub
```


