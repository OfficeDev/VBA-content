---
title: PicturePosition Property
keywords: fm20.chm5225079
f1_keywords:
- fm20.chm5225079
ms.prod: office
api_name:
- Office.PicturePosition
ms.assetid: dee7f263-90a9-cdeb-981f-65dd5e118a18
ms.date: 06/08/2017
---


# PicturePosition Property



Specifies the location of the picture relative to its caption.
 **Syntax**
 _object_. **PicturePosition** [= _fmPicturePosition_ ]
The  **PicturePosition** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmPicturePosition_|Optional. How the picture aligns with its container.|
 **Settings**
The settings for  _fmPicturePosition_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmPicturePositionLeftTop_|0|The picture appears to the left of the caption. The caption is aligned with the top of the picture.|
| _fmPicturePositionLeftCenter_|1|The picture appears to the left of the caption. The caption is centered relative to the picture.|
| _fmPicturePositionLeftBottom_|2|The picture appears to the left of the caption. The caption is aligned with the bottom of the picture.|
| _fmPicturePositionRightTop_|3|The picture appears to the right of the caption. The caption is aligned with the top of the picture.|
| _fmPicturePositionRightCenter_|4|The picture appears to the right of the caption. The caption is centered relative to the picture.|
| _fmPicturePositionRightBottom_|5|The picture appears to the right of the caption. The caption is aligned with the bottom of the picture.|
| _fmPicturePositionAboveLeft_|6|The picture appears above the caption. The caption is aligned with the left edge of the picture.|
| _fmPicturePositionAboveCenter_|7|The picture appears above the caption. The caption is centered below the picture (default).|
| _fmPicturePositionAboveRight_|8|The picture appears above the caption. The caption is aligned with the right edge of the picture.|
| _fmPicturePositionBelowLeft_|9|The picture appears below the caption. The caption is aligned with the left edge of the picture.|
| _fmPicturePositionBelowCenter_|10|The picture appears below the caption. The caption is centered above the picture.|
| _fmPicturePositionBelowRight_|11|The picture appears below the caption. The caption is aligned with the right edge of the picture.|
| _fmPicturePositionCenter_|12|The picture appears in the center of the control. The caption is centered horizontally and vertically on top of the picture.|
 **Remarks**
The picture and the caption, as a unit, are centered on the control. If no caption exists, the picture's location is relative to the center of the control.
This property is ignored if the  **Picture** property does not specify a picture.

