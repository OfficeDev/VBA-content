---
title: Shape.CopyPicture Method (Excel)
keywords: vbaxl10.chm636127
f1_keywords:
- vbaxl10.chm636127
ms.prod: excel
api_name:
- Excel.Shape.CopyPicture
ms.assetid: 276cd993-18b1-8c5b-3618-95e5b5c9a773
ms.date: 06/08/2017
---


# Shape.CopyPicture Method (Excel)

Copies the selected object to the Clipboard as a picture.


## Syntax

 _expression_ . **CopyPicture**( **_Appearance_** , **_Format_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Appearance_|Optional| **Variant**|A [XlPictureAppearance](xlpictureappearance-enumeration-excel.md) constant that specifies how the picture should be copied. The default value is **xlScreen** .|
| _Format_|Optional| **Variant**|A [XlCopyPictureFormat](xlcopypictureformat-enumeration-excel.md) constant that specifies the format of the picture. The default value is **xlPicture** .|

## Remarks

If you copy a range, it must be made up of adjacent cells.


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

