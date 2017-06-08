---
title: Hyperlink.ExtraInfo Property (Visio)
keywords: vis_sdr.chm15013490
f1_keywords:
- vis_sdr.chm15013490
ms.prod: visio
api_name:
- Visio.Hyperlink.ExtraInfo
ms.assetid: b5370912-5580-4c76-088d-265f87d1b37d
ms.date: 06/08/2017
---


# Hyperlink.ExtraInfo Property (Visio)

Returns or sets extra URL request information used to resolve the hyperlink's URL. Read/write.


## Syntax

 _expression_ . **ExtraInfo**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

String


## Remarks

Setting the  **ExtraInfo** property of a shape's **Hyperlink** object is optional, and is equivalent to setting the value of the ExtraInfo cell in the shape's Hyperlink. _name_ row.

You might, for example, set the  **Hyperlink** object's **ExtraInfo** property to the coordinates of an image map, the contents of a form, or a file name.

If the  **ExtraInfo** property you provide contains reserved characters other than spaces, you must input the escape character "%" and the character's hexidecimal equivalent. For example:

For "NAME=John Smith", set the  **ExtraInfo** property to "NAME=John Smith," because the extra information contains spaces, but no reserved characters.

For "PATH= _driveletter_ :\ _foldename_ ", set the **ExtraInfo** property to "PATH= _driveletter_ %3A%5C _foldername_ ", because of the reserved characters.


