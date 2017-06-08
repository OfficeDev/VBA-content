---
title: Miscellaneous Constants
keywords: vblr6.chm1092146
f1_keywords:
- vblr6.chm1092146
ms.prod: office
ms.assetid: ef7f52d8-5707-c7db-ca47-e7eaec37276d
ms.date: 06/08/2017
---


# Miscellaneous Constants

The following [constants](vbe-glossary.md) are defined in the Visual Basic for Applications[type library](vbe-glossary.md) and can be used anywhere in your code in place of the actual values:



|**Constant**|**Equivalent**|**Description**|
|:-----|:-----|:-----|
|**vbCrLf**|**Chr(** 13 **) + Chr(** 10 **)**|Carriage return-linefeed combination|
|**vbCr**|**Chr(** 13 **)**|Carriage return character|
|**vbLf**|**Chr(** 10 **)**|Linefeed character|
|**vbNewLine**|**Chr(** 13 **) + Chr(** 10 **)** or, on the Macintosh, **Chr(** 13 **)**|Platform-specific new line character; whichever is appropriate for current platform|
|**vbNullChar**|**Chr(** 0 **)**|Character having value 0|
|**vbNullString**|String having value 0|Not the same as a zero-length string (""); used for calling external procedures|
|**vbObjectError**|-2147221504|User-defined error numbers should be greater than this value. For example:
 `Err.Raise Number = vbObjectError + 1000`|
|**vbTab**|**Chr(** 9 **)**|Tab character|
|**vbBack**|**Chr(** 8 **)**|Backspace character|
|**vbFormFeed**|**Chr(** 12 **)**|Not useful in Microsoft Windows or on the Macintosh|
|**vbVerticalTab**|**Chr(** 11 **)**|Not useful in Microsoft Windows or on the Macintosh|

