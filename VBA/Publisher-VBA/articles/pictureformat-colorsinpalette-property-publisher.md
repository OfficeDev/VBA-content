---
title: PictureFormat.ColorsInPalette Property (Publisher)
keywords: vbapb10.chm3604754
f1_keywords:
- vbapb10.chm3604754
ms.prod: publisher
api_name:
- Publisher.PictureFormat.ColorsInPalette
ms.assetid: 34e671b1-af0e-0dac-1429-246facae975b
ms.date: 06/08/2017
---


# PictureFormat.ColorsInPalette Property (Publisher)

 Returns a **Long** that represents the number of colors in the picture's palette. Read-only.


## Syntax

 _expression_. **ColorsInPalette**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Long


## Remarks

This property applies only to pictures that are not TrueColor (that is, pictures that contain color data of less than 24 bits per channel). Returns "Permission Denied" for shapes representing pictures that are TrueColor.

Use the  **[IsTrueColor](pictureformat-istruecolor-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object to determine whether a picture contains color data of 24 bits per channel or greater.


## Example

The following example tests each picture in the active document and prints whether the picture is TrueColor. If the picture is not TrueColor, the example prints how many colors are in the picture's palette.


```vb
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 Debug.Print .Filename 
 If .IsTrueColor = msoTrue Then 
 Debug.Print "This picture is TrueColor" 
 Else 
 Debug.Print "This picture contains " &; .ColorsInPalette &; " colors." 
 End If 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 

```


