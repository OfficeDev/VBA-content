---
title: PictureFormat.ReplaceEx Method (Publisher)
keywords: vbapb10.chm3604808
f1_keywords:
- vbapb10.chm3604808
ms.prod: publisher
api_name:
- Publisher.PictureFormat.ReplaceEx
ms.assetid: 0f1b9eaf-51b6-ae21-518f-55663184ab87
ms.date: 06/08/2017
---


# PictureFormat.ReplaceEx Method (Publisher)

Replaces the specified picture, optionally fitting the replacement picture to the frame or filling the frame. Returns nothing.


## Syntax

 _expression_. **ReplaceEx**( _Pathname_,  _InsertAs_,  _Fit_)

 _expression_A variable that represents a  [PictureFormat](pictureformat-object-publisher.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Pathname|Required| **String**|The name and path of the file with which you want to replace the specified picture.|
|InsertAs|Optional| **[PbPictureInsertAs](pbpictureinsertas-enumeration-publisher.md)**|The manner in which you want the picture file inserted into the document: linked or embedded.|
|Fit|Optional| **[pbPictureInsertFit](pbpictureinsertfit-enumeration-publisher.md)**|Whether the inserted picture is fit to the frame, or fills the frame.|

## Remarks

The  _InsertAs_ parameter can be one of the following **PbPictureInsertAs** constants declared in the Microsoft Publisher type library. The default value is **pbPictureInsertAsOriginalState**.



| **pbPictureInsertAsEmbedded**|
| **pbPictureInsertAsLinked**|
| **pbPictureInsertAsOriginalState**|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ReplaceEx** method to replace all the pictures in a publication with a different picture. In this example, the replacement picture is fit to the frames of the previous pictures, but you can use **pbFill** in place of **pbFit** if you want to fill the frames instead. This example also excludes pictures on master pages.

Before running this macro, replace replacementPicturePath with the path to the picture you want to use as the replacement.




```vb
Public Sub ReplaceEx_Example()
    
    Dim pubPage As Page
    Dim pubShape As Shape
    Dim strReplacePicturePath As String
    
    strReplacePicturePath = replacementPicturePath
    
    For Each pubPage In ActiveDocument.Pages
        
        For Each pubShape In pubPage.Shapes
            
            If pubShape.Type = pbPicture Then

                pubShape.PictureFormat.ReplaceEx strReplacePicturePath, pbPictureInsertAsOriginalState, pbFit

            End If
        
        Next pubShape
        
    Next pubPage
            
End Sub
```


