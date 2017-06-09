---
title: Options.PathForPictures Property (Publisher)
keywords: vbapb10.chm1048596
f1_keywords:
- vbapb10.chm1048596
ms.prod: publisher
api_name:
- Publisher.Options.PathForPictures
ms.assetid: e66c8c86-f049-0f32-0a0d-60fd37470708
ms.date: 06/08/2017
---


# Options.PathForPictures Property (Publisher)

Returns a  **String** that represents the default path for picture files. Read.


## Syntax

 _expression_. **PathForPictures**

 _expression_A variable that represents a  **Options** object.


### Return Value

String


## Example

This example places the default path for picture files in a string and then uses the path string to add the specified file to the active publication. (Note that Filename must be replaced with a valid file name for this example to work.)


```vb
Sub InsertNewPicture() 
 Dim strPicPath As String 
 
 strPicPath = Options.PathForPictures 
 
 ActiveDocument.Pages(1).Shapes.AddPicture FileName:=strPicPath _ 
 &; "Filename", LinktoFile:=msoFalse, _ 
 SaveWithDocument:=msoTrue, Left:=50, Top:=50, Height:=200 
 
End Sub
```


