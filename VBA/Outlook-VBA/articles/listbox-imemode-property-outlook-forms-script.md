---
title: ListBox.IMEMode Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: c4e863d8-a581-2c45-92cc-1f6304692f76
ms.date: 06/08/2017
---


# ListBox.IMEMode Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the default run-time mode of the Input Method Editor (IME) for a control. Read/write.


## Syntax

 _expression_. **IMEMode**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

This property applies only to applications written for Asian languages and is ignored in other applications.

The settings for fmIMEMode are:



|**Value**|**Description**|
|:-----|:-----|
|0|Does not control IME (default).|
|1|IME on.|
|2|IME off. English mode.|
|3|IME off. User can't turn on IME by keyboard.|
|4|IME on with Full-width Hiragana mode.|
|5|IME on with Full-width Katakana mode.|
|6|IME on with Half-width Katakana mode.|
|7|IME on with Full-width Alphanumeric mode.|
|8|IME on with Half-width Alphanumeric mode.|
|9|IME on with Full-width Hangul mode.|
|10|IME on with Half-width Hangul mode.|
A setting of 0 indicates that the mode of the IME does not change when the control receives focus at run time. For any other value, the mode of the IME is set to the value specified by the  **IMEMode** property when the control receives focus at run time.


