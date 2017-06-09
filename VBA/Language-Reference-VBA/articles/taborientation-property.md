---
title: TabOrientation Property
keywords: fm20.chm5225100
f1_keywords:
- fm20.chm5225100
ms.prod: office
api_name:
- Office.TabOrientation
ms.assetid: dc84899d-2c50-56d2-5178-f8bfaefaa165
ms.date: 06/08/2017
---


# TabOrientation Property



Specifies the location of the tabs on a  **MultiPage** or **TabStrip**.
 **Syntax**
 _object_. **TabOrientation** [= _fmTabOrientation_ ]
The  **TabOrientation** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmTabOrientation_|Optional. Where the tabs will appear.|
 **Settings**
The settings for  _fmTabOrientation_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmTabOrientationTop_|0|The tabs appear at the top of the control (default).|
| _fmTabOrientationBottom_|1|The tabs appear at the bottom of the control.|
| _fmTabOrientationLeft_|2|The tabs appear at the left side of the control.|
| _fmTabOrientationRight_|3|The tabs appear at the right side of the control.|
 **Remarks**
If you use TrueType fonts, the text rotates when the  **TabOrientation** property is set to **fmTabOrientationLeft** or **fmTabOrientationRight**. If you use bitmapped fonts, the text does not rotate.

