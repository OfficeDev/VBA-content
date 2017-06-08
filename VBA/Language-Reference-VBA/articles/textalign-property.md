---
title: TextAlign Property
keywords: fm20.chm5225104
f1_keywords:
- fm20.chm5225104
ms.prod: office
api_name:
- Office.TextAlign
ms.assetid: 31904bca-6238-6807-fdbd-463cbc82b8ed
ms.date: 06/08/2017
---


# TextAlign Property



Specifies how text is aligned in a control.
 **Syntax**
 _object_. **TextAlign** [= _fmTextAlign_ ]
The  **TextAlign** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmTextAlign_|Optional. How text is aligned in the control.|
 **Settings**
The settings for  _fmTextAlign_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmTextAlignLeft_|1|Aligns the first character of displayed text with the left edge of the control's display or edit area (default).|
| _fmTextAlignCenter_|2|Centers the text in the control's display or edit area.|
| _fmTextAlignRight_|3|Aligns the last character of displayed text with the right edge of the control's display or edit area.|
 **Remarks**
For a  **ComboBox**, the **TextAlign** property only affects the edit region; this property has no effect on the alignment of text in the list. For stand-alone labels, **TextAlign** determines the alignment of the label's caption.

