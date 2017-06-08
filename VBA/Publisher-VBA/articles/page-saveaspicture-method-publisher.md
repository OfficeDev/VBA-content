---
title: Page.SaveAsPicture Method (Publisher)
keywords: vbapb10.chm393272
f1_keywords:
- vbapb10.chm393272
ms.prod: publisher
api_name:
- Publisher.Page.SaveAsPicture
ms.assetid: 9b118126-e072-9516-9863-14ea60264f01
ms.date: 06/08/2017
---


# Page.SaveAsPicture Method (Publisher)

Saves a page as a picture file.


## Syntax

 _expression_. **SaveAsPicture**( **_Filename_**,  **_pbResolution_**)

 _expression_A variable that represents a  **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Filename|Required| **String**|The path and file name of the new picture file you want to create. The graphics format the picture is saved in is determined by the file name extension (such as .jpg or .gif) you specify.|
|pbResolution|Optional| **PbPictureResolution**|The resolution in which you want the picture to be saved. See Remarks for possible values.|

## Remarks

Possible values for the pbResolution parameter are declared in the  **[PbPictureResolution](pbpictureresolution-enumeration-publisher.md)** enumeration in the Microsoft Publisher type library.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SaveAsPicture** method to save the first page of the active publication as a .jpg picture file.

Before running this code, replace  _filename.jpg_ with a valid file name and the path to a folder on your computer where you have permission to save files.




```vb
Public Sub SaveAsPicture_Example() 
 
 ThisDocument.Pages(1).SaveAsPicture "filename.jpg" 
 
End Sub
```


