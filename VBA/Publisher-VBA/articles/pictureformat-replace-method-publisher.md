---
title: PictureFormat.Replace Method (Publisher)
keywords: vbapb10.chm3604786
f1_keywords:
- vbapb10.chm3604786
ms.prod: publisher
api_name:
- Publisher.PictureFormat.Replace
ms.assetid: b2bce79a-5c46-1473-601d-a4a25176edeb
ms.date: 06/08/2017
---


# PictureFormat.Replace Method (Publisher)

Replaces the specified picture. Returns  **Nothing**.


## Syntax

 _expression_. **Replace**( **_Pathname_**,  **_InsertAs_**)

 _expression_A variable that represents a  **PictureFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Pathname|Required| **String**|The name and path of the file with which you want to replace the specified picture.|
|InsertAs|Optional| **PbPictureInsertAs**|The manner in which you want the picture file inserted into the document: linked or embedded.|

## Remarks

Use the  **Replace** method to update linked picture files that have been modified since they were inserted into the document. Use the **[LinkedFileStatus](pictureformat-linkedfilestatus-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object to determine if a linked picture has been modified.

The InsertAs parameter can be one of the following  **PbPictureInsertAs** constants declared in the Microsoft Publisher type library. the default value is **pbPictureInsertAsOriginalState**.



| **pbPictureInsertAsEmbedded**|
| **pbPictureInsertAsLinked**|
| **pbPictureInsertAsOriginalState**|

## Example

The following example replaces every occurrence of a specific picture in the active publication with another picture.


```vb
Sub ReplaceLogo() 
 
Dim pgLoop As Page 
Dim shpLoop As Shape 
Dim strExistingArtName As String 
Dim strReplaceArtName As String 
 
 
strExistingArtName = "C:\path\logo 1.bmp" 
strReplaceArtName = "C:\path\logo 2.bmp" 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .Filename = strExistingArtName Then 
 .Replace (strReplaceArtName) 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
Next pgLoop 
 
End Sub
```

This example tests each linked picture to determine if the linked file has been modified since it was inserted into the publication. If it has, the picture is updated by replacing the file with itself.




```vb
Sub UpdateModifiedLinkedPictures() 
 
Dim pgLoop As Page 
Dim shpLoop As Shape 
Dim strPictureName As String 
 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .LinkedFileStatus = pbLinkedFileModified Then 
 strPictureName = .Filename 
 .Replace (strPictureName) 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 
 
End Sub
```


