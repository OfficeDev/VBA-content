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


| <strong>Part</strong>     | <strong>Description</strong>          |
|:--------------------------|:--------------------------------------|
| <em>object</em>           | Required. A valid object.             |
| <em>fmTabOrientation</em> | Optional. Where the tabs will appear. |

 **Settings**
The settings for  _fmTabOrientation_ are:


| <strong>Constant</strong>       | <strong>Value</strong> | <strong>Description</strong>                         |
|:--------------------------------|:-----------------------|:-----------------------------------------------------|
| <em>fmTabOrientationTop</em>    | 0                      | The tabs appear at the top of the control (default). |
| <em>fmTabOrientationBottom</em> | 1                      | The tabs appear at the bottom of the control.        |
| <em>fmTabOrientationLeft</em>   | 2                      | The tabs appear at the left side of the control.     |
| <em>fmTabOrientationRight</em>  | 3                      | The tabs appear at the right side of the control.    |

 **Remarks**
If you use TrueType fonts, the text rotates when the  **TabOrientation** property is set to **fmTabOrientationLeft** or **fmTabOrientationRight**. If you use bitmapped fonts, the text does not rotate.

