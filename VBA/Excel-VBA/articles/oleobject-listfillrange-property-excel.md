---
title: OLEObject.ListFillRange Property (Excel)
keywords: vbaxl10.chm417082
f1_keywords:
- vbaxl10.chm417082
ms.prod: excel
api_name:
- Excel.OLEObject.ListFillRange
ms.assetid: d8a44f9f-49bb-237b-66c8-9f6c06fe82ac
ms.date: 06/08/2017
---


# OLEObject.ListFillRange Property (Excel)

Returns or sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box. Read/write  **String** .


## Syntax

 _expression_ . **ListFillRange**

 _expression_ A variable that represents an **OLEObject** object.


## Remarks

Microsoft Excel reads the contents of every cell in the range and inserts the cell values into the list box. The list tracks changes in the range's cells.

If the list in the list box was created with the  **[AddItem](controlformat-additem-method-excel.md)** method, this property returns an empty string ("").


## See also


#### Concepts


[OLEObject Object](oleobject-object-excel.md)

