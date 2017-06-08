---
title: UserPicture Method
keywords: vbagr10.chm67165
f1_keywords:
- vbagr10.chm67165
ms.prod: excel
api_name:
- Excel.UserPicture
ms.assetid: ad8e3079-c063-2bb6-e462-11a0e8ecfba6
ms.date: 06/08/2017
---


# UserPicture Method

Fills the specified shape with an image.

 _expression_. **UserPicture**( **_PictureFile_**,  **_PictureFormat_**,  **_PictureStackUnit_**,  **_PicturePlacement_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **PictureFile** Required **Variant**. The name of the specified picture file.
 **PictureFormat** Optional
 **XlChartPictureType**
. The format of the specified picture.


|XlChartPictureType can be one of these XlChartPictureType constants.|
| **xlScale**|
| **xlStackScale** **xlStack** **xlStretch**|
 **PictureStackUnit** Optional **Variant**. The stack or scale unit for the specified picture (depends on the  **_PictureFormat_** argument).
 **PicturePlacement** Optional
 **XlChartPicturePlacement**
. The placement of the specified picture.


|XlChartPicturePlacement can be one of these XlChartPicturePlacement constants.|
| **xlSides**|
| **xlEnd** **xlEndSides** **xlFront** **xlFrontSides** **xlFrontEnd** **xlAllFaces**|

## Example

This example sets the chart's fill format so that it's based on a user-supplied picture.


```vb
With myChart.ChartArea.Fill 
 .UserPicture PictureFile:="C:\My Documents\brick.bmp" 
 .Visible = True 
End With
```


