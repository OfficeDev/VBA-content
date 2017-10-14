---
title: PictureFormat.FileSize Property (Publisher)
keywords: vbapb10.chm3604757
f1_keywords:
- vbapb10.chm3604757
ms.prod: publisher
api_name:
- Publisher.PictureFormat.FileSize
ms.assetid: 8bad7bc0-7381-9bd8-3db8-5841e41ccb34
ms.date: 06/08/2017
---


# PictureFormat.FileSize Property (Publisher)

Returns a  **Long** that represents, in bytes, the size of the picture or OLE object as it appears in the specified publication. Read-only.


## Syntax

 _expression_. **FileSize**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Long


## Remarks

If the picture or OLE object is linked, use the  **[OriginalFileSize](pictureformat-originalfilesize-property-publisher.md)** property to determine the size of the linked file.

To determine whether a shape represents a linked picture, use either the  **[Type](shape-type-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object, or the **[IsLinked](pictureformat-islinked-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object.


## Example

The following example tests each picture in the active publication, and prints selected image properties for pictures that are linked.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Original File Size: " &; .OriginalFileSize &; " bytes" 
 Debug.Print "File size in publication: " &; .FileSize &; " bytes" 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


