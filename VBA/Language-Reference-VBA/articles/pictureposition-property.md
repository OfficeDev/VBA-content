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


| <strong>Part</strong>      | <strong>Description</strong>                         |
|:---------------------------|:-----------------------------------------------------|
| <em>object</em>            | Required. A valid object.                            |
| <em>fmPicturePosition</em> | Optional. How the picture aligns with its container. |

 **Settings**
The settings for  _fmPicturePosition_ are:


| <strong>Constant</strong>             | <strong>Value</strong> | <strong>Description</strong>                                                                                                 |
|:--------------------------------------|:-----------------------|:-----------------------------------------------------------------------------------------------------------------------------|
| <em>fmPicturePositionLeftTop</em>     | 0                      | The picture appears to the left of the caption. The caption is aligned with the top of the picture.                          |
| <em>fmPicturePositionLeftCenter</em>  | 1                      | The picture appears to the left of the caption. The caption is centered relative to the picture.                             |
| <em>fmPicturePositionLeftBottom</em>  | 2                      | The picture appears to the left of the caption. The caption is aligned with the bottom of the picture.                       |
| <em>fmPicturePositionRightTop</em>    | 3                      | The picture appears to the right of the caption. The caption is aligned with the top of the picture.                         |
| <em>fmPicturePositionRightCenter</em> | 4                      | The picture appears to the right of the caption. The caption is centered relative to the picture.                            |
| <em>fmPicturePositionRightBottom</em> | 5                      | The picture appears to the right of the caption. The caption is aligned with the bottom of the picture.                      |
| <em>fmPicturePositionAboveLeft</em>   | 6                      | The picture appears above the caption. The caption is aligned with the left edge of the picture.                             |
| <em>fmPicturePositionAboveCenter</em> | 7                      | The picture appears above the caption. The caption is centered below the picture (default).                                  |
| <em>fmPicturePositionAboveRight</em>  | 8                      | The picture appears above the caption. The caption is aligned with the right edge of the picture.                            |
| <em>fmPicturePositionBelowLeft</em>   | 9                      | The picture appears below the caption. The caption is aligned with the left edge of the picture.                             |
| <em>fmPicturePositionBelowCenter</em> | 10                     | The picture appears below the caption. The caption is centered above the picture.                                            |
| <em>fmPicturePositionBelowRight</em>  | 11                     | The picture appears below the caption. The caption is aligned with the right edge of the picture.                            |
| <em>fmPicturePositionCenter</em>      | 12                     | The picture appears in the center of the control. The caption is centered horizontally and vertically on top of the picture. |

 **Remarks**
The picture and the caption, as a unit, are centered on the control. If no caption exists, the picture's location is relative to the center of the control.
This property is ignored if the  **Picture** property does not specify a picture.

