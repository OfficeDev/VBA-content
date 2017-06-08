---
title: PictureFormat.EffectiveResolution Property (Publisher)
keywords: vbapb10.chm3604755
f1_keywords:
- vbapb10.chm3604755
ms.prod: publisher
api_name:
- Publisher.PictureFormat.EffectiveResolution
ms.assetid: 33e5323f-5e10-b2ed-62eb-03ecbbb1e893
ms.date: 06/08/2017
---


# PictureFormat.EffectiveResolution Property (Publisher)

Returns a  **Long** that represents, in dots per inch (dpi), the effective resolution of the picture. Read-only.


## Syntax

 _expression_. **EffectiveResolution**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

Long


## Remarks

The effective resolution of a picture is inversely proportional to the scaling at which the picture is printed. The larger the scaling, the lower the effective resolution. For example, suppose a picture measuring 4 inches by 4 inches was originally scanned at 300 dpi. If that picture is scaled to 2 inches by 2 inches, its effective resolution is 600 dpi.

Use the  **[OriginalResolution](pictureformat-originalresolution-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object to determine the resolution of linked pictures or OLE objects. Use the **[HorizontalScale](pictureformat-horizontalscale-property-publisher.md)** and **[VerticalScale](pictureformat-verticalscale-property-publisher.md)** properties to determine the scaling of a picture.


## Example

The following example returns a list of pictures whose effective resolution falls below a specified threshold (100 dpi) in the active publication.


```vb
Sub ListLowResolutionPictures() 
 Dim pgLoop As Page 
 Dim shpLoop As Shape 
 
 For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 If .EffectiveResolution < 100 Then 
 Debug.Print .Filename 
 Debug.Print "Page " &; pgLoop.PageNumber 
 Debug.Print "Resolution in publication: " &; .EffectiveResolution 
 End If 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
 Next pgLoop 
 
End Sub
```


