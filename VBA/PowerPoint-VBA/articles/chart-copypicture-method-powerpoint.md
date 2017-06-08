---
title: Chart.CopyPicture Method (PowerPoint)
keywords: vbapp10.chm684022
f1_keywords:
- vbapp10.chm684022
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.CopyPicture
ms.assetid: ac8c3f05-3458-8f24-ada8-b89beb52a968
ms.date: 06/08/2017
---


# Chart.CopyPicture Method (PowerPoint)

Copies the selected object to the Clipboard as a picture.


## Syntax

 _expression_. **CopyPicture**( **_Appearance_**, **_Format_**, **_Size_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Appearance_|Optional|**[XlPictureAppearance](xlpictureappearance-enumeration-powerpoint.md)**|One of the enumeration values that specifies how the picture should be copied. The default is  **xlScreen**.|
| _Format_|Optional|**[XlCopyPictureFormat](xlcopypictureformat-enumeration-powerpoint.md)**|One of the enumeration values that specifies the format of the picture. The default is  **xlPicture**.|
| _Size_|Optional|**XlPictureAppearance**|One of the enumeration values that specifies the size of the copied picture when the object is a chart on a chart sheet (not embedded on a worksheet). The default is  **xlPrinter**.|

## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

