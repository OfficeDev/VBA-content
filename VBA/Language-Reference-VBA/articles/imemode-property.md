---
title: IMEMode Property
keywords: fm20.chm5225043
f1_keywords:
- fm20.chm5225043
ms.prod: office
api_name:
- Office.IMEMode
ms.assetid: b47dd67c-f058-ad85-97ae-17efe46875ed
ms.date: 06/08/2017
---


# IMEMode Property



Specifies the default [run time](vbe-glossary.md) mode of the[Input Method Editor (IME](glossary-vba.md)) for a control. This property applies only to applications written for East Asia and is ignored in other applications.
 **Syntax**
 _object_. **IMEMode** [= _fmIMEMode_ ]
The  **IMEMode** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmIMEMode_|Optional. The mode of the Input Method Editor (IME).|
 **Settings**
The settings for  _fmIMEMode_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmIMEModeNoControl_|0|Does not control IME (default).|
| _fmIMEModeOn_|1|IME on.|
| _fmIMEModeOff_|2|IME off. English mode.|
| _fmIMEModeDisable_|3|IME off. User can't turn on IME by keyboard.|
| _fmIMEModeHiragana_|4|IME on with Full-width Hiragana mode.|
| _fmIMEModeKatakana_|5|IME on with Full-width Katakana mode.|
| _fmIMEModeKatakanaHalf_|6|IME on with Half-width Katakana mode.|
| _fmIMEModeAlphaFull_|7|IME on with Full-width Alphanumeric mode.|
| _fmIMEModeAlpha_|8|IME on with Half-width Alphanumeric mode.|
| _fmIMEModeHangulFull_|9|IME on with Full-width Hangul mode.|
| _fmIMEModeHangul_|10|IME on with Half-width Hangul mode.|
The  **fmIMEModeNoControl** setting indicates that the mode of the IME does not change when the control receives[focus](vbe-glossary.md) at run time. For any other value, the mode of the IME is set to the value specified by the **IMEMode** property when the control receives focus at run time.
 **Remarks**
There are two ways to set the mode of the IME. One is through the toolbar of the IME. The other is with the  **IMEMode** property of a control, which sets or returns the current mode of the IME. This property allows dynamic control of the IME through code.
The following example explains how  **IMEMode** interacts with the toolbar of the IME. Assume that you have designed a form with TextBox1 and CheckBox1. You have set TextBox1.IMEMode to 0, and you have set CheckBox1.IMEMode to 1. While in design mode you have used the IME toolbar to put the IME in mode 2.
When you run the form, the IME begins in mode 2. If you click TextBox1, the IME mode does not change because  **IMEMode** for this control is 0. If you click CheckBox1, the IME changes to mode 1, because **IMEMode** for this control is 1. If you click again on TextBox1, the IME remains in mode 1 ( **IMEMode** is 0, so the IME retains its last setting).
However, you can override  **IMEMode**. For example, assume you click CheckBox1 and the IME enters mode 1, as defined by **IMEMode** for the **CheckBox**. If you then use the IME toolbar to put the IME in mode 3, then the IME will be set to mode 3 anytime you click the control. This does not change the value of the property, it overrides the property until the next time you run the form.

