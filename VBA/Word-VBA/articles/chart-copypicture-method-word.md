---
title: Chart.CopyPicture Method (Word)
keywords: vbawd10.chm79364163
f1_keywords:
- vbawd10.chm79364163
ms.prod: word
api_name:
- Word.Chart.CopyPicture
ms.assetid: 90f41c1a-8a96-0959-6c9a-b10f7f4744b0
ms.date: 06/08/2017
---


# Chart.CopyPicture Method (Word)

Copies the selected object to the Clipboard as a picture.


## Syntax

 _expression_ . **CopyPicture**( **_Appearance_** , **_Format_** , **_Size_** )

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Appearance_|Optional| **[XlPictureAppearance](xlpictureappearance-enumeration-word.md)**|One of the enumeration values that specifies how the picture should be copied. The default is  **xlScreen** .|
| _Format_|Optional| **[XlCopyPictureFormat](xlcopypictureformat-enumeration-word.md)**|One of the enumeration values that specifies the format of the picture. The default is  **xlPicture** .|
| _Size_|Optional| **XlPictureAppearance**|One of the enumeration values that specifies the size of the copied picture when the object is a chart on a chart sheet (not embedded on a worksheet). The default is  **xlPrinter** .|

## See also


#### Concepts


[Chart Object](chart-object-word.md)

