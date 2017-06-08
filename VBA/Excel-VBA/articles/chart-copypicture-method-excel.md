---
title: Chart.CopyPicture Method (Excel)
keywords: vbaxl10.chm149095
f1_keywords:
- vbaxl10.chm149095
ms.prod: excel
api_name:
- Excel.Chart.CopyPicture
ms.assetid: f69451cd-4be5-982a-58b8-63e0f24e0261
ms.date: 06/08/2017
---


# Chart.CopyPicture Method (Excel)

Copies the selected object to the Clipboard as a picture.


## Syntax

 _expression_ . **CopyPicture**( **_Appearance_** , **_Format_** , **_Size_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Appearance_|Optional| **[XlPictureAppearance](xlpictureappearance-enumeration-excel.md)**|. Specifies how the picture should be copied. The default value is  **xlScreen** .|
| _Format_|Optional| **[XlCopyPictureFormat](xlcopypictureformat-enumeration-excel.md)**|. The format of the picture. The default value is  **xlPicture** .|
| _Size_|Optional| **XlPictureAppearance**|The size of the copied picture when the object is a chart on a chart sheet (not embedded on a worksheet). The default value is  **xlPrinter** .|

## See also


#### Concepts


[Chart Object](chart-object-excel.md)

