---
title: FormatNumber Function
keywords: vblr6.chm1008937
f1_keywords:
- vblr6.chm1008937
ms.prod: office
ms.assetid: ab4012b3-efed-bc06-9c5e-416c9200ffed
ms.date: 06/08/2017
---


# FormatNumber Function



 **Description**
Returns an expression formatted as a number.
 **Syntax**
 **FormatNumber(**_Expression_ [ **,**_NumDigitsAfterDecimal_ [ **,**_IncludeLeadingDigit_ [ **,**_UseParensForNegativeNumbers_ [ **,**_GroupDigits_ ]]]] **)**
The  **FormatNumber** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _Expression_|Required. Expression to be formatted.|
| _NumDigitsAfterDecimal_|Optional. Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used.|
| _IncludeLeadingDigit_|Optional. Tristate constant that indicates whether or not a leading zero is displayed for fractional values. See Settings section for values.|
| _UseParensForNegativeNumbers_|Optional. Tristate constant that indicates whether or not to place negative values within parentheses. See Settings section for values.|
| _GroupDigits_|Optional. Tristate constant that indicates whether or not numbers are grouped using the group delimiter specified in the computer's regional settings. See Settings section for values.|
 **Settings**
The  _IncludeLeadingDigit_, _UseParensForNegativeNumbers_, and _GroupDigits_ arguments have the following settings:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbTrue**|-1|True|
|**vbFalse**| 0|False|
|**vbUseDefault**|-2|Use the setting from the computer's regional settings.|
 **Remarks**
When one or more optional arguments are omitted, the values for omitted arguments are provided by the computer's regional settings.

 **Note**  All settings information comes from the  **Regional Settings Number** tab.


