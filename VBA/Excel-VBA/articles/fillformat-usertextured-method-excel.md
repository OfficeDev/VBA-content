---
title: FillFormat.UserTextured Method (Excel)
keywords: vbaxl10.chm115010
f1_keywords:
- vbaxl10.chm115010
ms.prod: excel
api_name:
- Excel.FillFormat.UserTextured
ms.assetid: 8c8e7569-50e9-fec5-9c0e-195b26f9394c
ms.date: 06/08/2017
---


# FillFormat.UserTextured Method (Excel)

Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the  **[UserPicture](fillformat-userpicture-method-excel.md)** method.


## Syntax

 _expression_ . **UserTextured**( **_TextureFile_** )

 _expression_ A variable that represents a **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TextureFile_|Required| **String**| The name of the picture file.|

## Example

This example sets the fill format for chart two.


```vb
Charts(2).ChartArea.Fill.UserTextured "brick.gif"
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

