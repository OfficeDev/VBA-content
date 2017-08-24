---
title: PictureFormat.HorizontalScale Property (Publisher)
keywords: vbapb10.chm3604760
f1_keywords:
- vbapb10.chm3604760
ms.prod: publisher
api_name:
- Publisher.PictureFormat.HorizontalScale
ms.assetid: 7be51cde-5b2d-7870-7f39-2fa9bd714d68
ms.date: 06/08/2017
---


# PictureFormat.HorizontalScale Property (Publisher)

Returns a  **Long** that represents the scaling of the picture along its horizontal axis. The scaling is expressed as a percentage (for example, 200 equals 200 percent scaling). Read-only.


## Syntax

 _expression_. **HorizontalScale**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Long


## Remarks

The effective resolution of a picture is inversely proportional to the scaling at which the picture is printed. The larger the scaling, the lower the effective resolution. For example, suppose a picture measuring 4 inches by 4 inches was originally scanned at 300 dpi. If that picture is scaled to 2 inches by 2 inches, its effective resolution is 600 dpi.

Use the  **[EffectiveResolution](pictureformat-effectiveresolution-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object to determine the resolution at which the picture or OLE object will print in the specified document.


## Example

The following example prints selected image properties for each picture in the active publication.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 If .IsEmpty = msoFalse Then 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Resolution in Publication: " &; .EffectiveResolution &; " dpi" 
 Debug.Print "Horizontal Scaling: " &; .HorizontalScale &; "%" 
 Debug.Print "Height in publication: " &; .Height &; " points" 
 Debug.Print "Vertical Scaling: " &; .VerticalScale &; "%" 
 Debug.Print "Width in publication: " &; .Width &; " points" 
 End If 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 
 

```


