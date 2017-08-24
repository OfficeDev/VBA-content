---
title: PictureFormat.OriginalWidth Property (Publisher)
keywords: vbapb10.chm3604777
f1_keywords:
- vbapb10.chm3604777
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalWidth
ms.assetid: 3c418f3f-b2af-3176-9a37-a548b15fb4bc
ms.date: 06/08/2017
---


# PictureFormat.OriginalWidth Property (Publisher)

Returns a  **Variant** that represents, in points, the width of the specified linked picture or OLE object. Read-only.


## Syntax

 _expression_. **OriginalWidth**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

Variant


## Remarks

This property applies only to linked pictures. Returns "Permission Denied" for shapes representing embedded or pasted pictures.

To determine whether a shape represents a linked picture, use either the  **[Type](shape-type-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object, or the **[IsLinked](pictureformat-islinked-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object.


## Example

The following example tests each picture in the active publication, and returns selected image properties for pictures that are linked.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Vertical Scaling: " &; .VerticalScale &; "%" 
 Debug.Print "Original Image Width: " &; .OriginalWidth &; " points" 
 Debug.Print "Width in publication: " &; .Width &; " points" 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


