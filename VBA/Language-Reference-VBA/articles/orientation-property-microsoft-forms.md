---
title: Orientation Property (Microsoft Forms)
keywords: fm20.chm5225074
f1_keywords:
- fm20.chm5225074
ms.prod: office
ms.assetid: 3e57f9af-8aa5-85f5-f3af-81f9a61373c0
ms.date: 06/08/2017
---


# Orientation Property (Microsoft Forms)



Specifies whether the  **SpinButton** or **ScrollBar** is oriented vertically or horizontally.
 **Syntax**
 _object_. **Orientation** [= _fmOrientation_ ]
The  **Orientation** property syntax has these parts:


| <strong>Part</strong>  | <strong>Description</strong>          |
|:-----------------------|:--------------------------------------|
| <em>object</em>        | Required. A valid object.             |
| <em>fmOrientation</em> | Optional. Orientation of the control. |

 **Settings**
The settings for  _fmOrientation_ are:


| <strong>Constant</strong>        | <strong>Value</strong> | <strong>Description</strong>                                                                 |
|:---------------------------------|:-----------------------|:---------------------------------------------------------------------------------------------|
| <em>fmOrientationAuto</em>       | -1                     | Automatically determines the orientation based upon the dimensions of the control (default). |
| <em>FmOrientationVertical</em>   | 0                      | Control is rendered vertically.                                                              |
| <em>FmOrientationHorizontal</em> | 1                      | Control is rendered horizontally.                                                            |

 **Remarks**
If you specify automatic orientation, the height and width of the control determine whether it appears horizontally or vertically. For example, if the control is wider than it is tall, it appears horizontally; if it is taller than it is wide, the control appears vertically.

