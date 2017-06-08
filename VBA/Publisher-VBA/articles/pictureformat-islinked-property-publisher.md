---
title: PictureFormat.IsLinked Property (Publisher)
keywords: vbapb10.chm3604769
f1_keywords:
- vbapb10.chm3604769
ms.prod: publisher
api_name:
- Publisher.PictureFormat.IsLinked
ms.assetid: 2215cee8-864d-7228-8692-a428385d2be2
ms.date: 06/08/2017
---


# PictureFormat.IsLinked Property (Publisher)

Returns a  **MsoTriState** constant indicating whether the specified picture is a linked picture or OLE object. Read-only.


## Syntax

 _expression_. **IsLinked**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

MsoTriState


## Remarks

Returns  **msoFalse** for pasted or embedded pictures and OLE objects.

If a picture or OLE object is linked, several additional properties of the  **[PictureFormat](pictureformat-object-publisher.md)** object dealing with the original picture (such as ** [OriginalFileSize](pictureformat-originalfilesize-property-publisher.md)**) are accessible.

The  **IsLinked** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The specified picture is not a linked picture.|
| **msoTriStateMixed**|Indicates a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The specified picture is a linked picture.|

## Example

The following example returns whether the first shape on the first page of the active publication contains an alpha channel. If the picture is linked, and the original picture contains an alpha channel, that is also returned. This example assumes the shape is a picture.


```vb
With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 If .HasAlphaChannel = msoTrue Then 
 Debug.Print .Filename 
 Debug.Print "This picture contains an alpha channel." 
 
 If .IsLinked = msoTrue Then 
 If .OriginalHasAlphaChannel = msoTrue Then 
 Debug.Print "The linked picture " &; _ 
 "also contains an alpha channel." 
 End If 
 End If 
 End If 
End With 

```


