---
title: Page.PictureType Property (Access)
keywords: vbaac10.chm12150
f1_keywords:
- vbaac10.chm12150
ms.prod: access
api_name:
- Access.Page.PictureType
ms.assetid: 113e63ec-a2d9-bd5e-8a49-82fd92339cb8
ms.date: 06/08/2017
---


# Page.PictureType Property (Access)

You can use the  **PictureType** property to specify whether Microsoft Access stores an object's picture as a linked or an embedded object. Read/write **Byte**.


## Syntax

 _expression_. **PictureType**

 _expression_ A variable that represents a **Page** object.


## Remarks

The  **PictureType** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|(Default) The picture is embedded in the object and becomes part of the database file.|
|1|The picture is linked to the object. Microsoft Access stores a pointer to the location of the picture on the disk.|
This property can be set only in form Design view or report Design view.

For controls, you can set the default for this property by using the default control style or the  **DefaultControl** property in Visual Basic.

When this property is set to 0, the size of the database increases by the size of the picture file and, with some .wmf files, the size may increase as much as twice the size of the picture file. When this property is set to 1, there is no increase in the size of the database because Microsoft Access only saves a pointer to the picture's location on the disk.


 **Note**   If a linked file is moved to another location on the disk, you must re-establish the link by using the object's **Picture** property.

For embedded pictures, the object's  **PictureData** property stores the individual bits that make up a picture's image. For linked pictures, this property stores the path to the picture's file.

You can modify a linked picture by using a separate application and changes to the picture will appear the next time the object containing that picture is displayed in the database.


## Example

The following example links the picture in the "Customer Photo" image control on the "Order Entry" form to the picture on the disk. The picture is not part of the database file.


```vb
Forms("Order Entry").Controls("Customer Photo").PictureType = 1 

```


## See also


#### Concepts


[Page Object](page-object-access.md)

