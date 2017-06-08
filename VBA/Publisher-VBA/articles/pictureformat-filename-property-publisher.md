---
title: PictureFormat.Filename Property (Publisher)
keywords: vbapb10.chm3604756
f1_keywords:
- vbapb10.chm3604756
ms.prod: publisher
api_name:
- Publisher.PictureFormat.Filename
ms.assetid: 73e2a224-f15a-50cc-462e-10ccf9478122
ms.date: 06/08/2017
---


# PictureFormat.Filename Property (Publisher)

Returns a  **String** that represents the file name of the specified picture or OLE object. Read-only.


## Syntax

 _expression_. **Filename**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

String


## Remarks

For linked pictures and OLE objects, the returned string represents the full path and file name of the picture. For embedded pictures and OLE objects, the returned string represents the file name only.

To determine whether a shape represents a linked picture, use either the  **[Type](shape-type-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object, or the **[IsLinked](pictureformat-islinked-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object.


## Example

The following example returns selected image properties for each picture in the active publication.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 If .IsEmpty = msoFalse Then 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Horizontal Scaling: " &; .HorizontalScale &; "%" 
 Debug.Print "Vertical Scaling: " &; .VerticalScale &; "%" 
 Debug.Print "File size in publication: " &; .FileSize &; " bytes" 
 
 End If 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop
```


