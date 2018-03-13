---
title: FormatPercent Function
keywords: vblr6.chm1008936
f1_keywords:
- vblr6.chm1008936
ms.prod: office
ms.assetid: 46da0c9c-f6a7-9349-7d5f-fc8147e7a6b3
ms.date: 06/08/2017
---


# FormatPercent Function



 **Description**
Returns an expression formatted as a percentage (multipled by 100) with a trailing % character.
 **Syntax**
 **FormatPercent(**_Expression_ [ **,**_NumDigitsAfterDecimal_ [ **,**_IncludeLeadingDigit_ [ **,**_UseParensForNegativeNumbers_ [ **,**_GroupDigits_ ]]]] **)**
The  **FormatPercent** function syntax has these parts:


| <strong>Part</strong>                | <strong>Description</strong>                                                                                                                                                            |
|:-------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>Expression</em>                  | Required. Expression to be formatted.                                                                                                                                                   |
| <em>NumDigitsAfterDecimal</em>       | Optional. Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used.      |
| <em>IncludeLeadingDigit</em>         | Optional. Tristate constant that indicates whether or not a leading zero is displayed for fractional values. See Settings section for values.                                           |
| <em>UseParensForNegativeNumbers</em> | Optional. Tristate constant that indicates whether or not to place negative values within parentheses. See Settings section for values.                                                 |
| <em>GroupDigits</em>                 | Optional. Tristate constant that indicates whether or not numbers are grouped using the group delimiter specified in the computer's regional settings. See Settings section for values. |

 **Settings**
The  _IncludeLeadingDigit_, _UseParensForNegativeNumbers_, and _GroupDigits_ arguments have the following settings:


| <strong>Constant</strong>     | <strong>Value</strong> | <strong>Description</strong>                           |
|:------------------------------|:-----------------------|:-------------------------------------------------------|
| <strong>vbTrue</strong>       | -1                     | True                                                   |
| <strong>vbFalse</strong>      | 0                      | False                                                  |
| <strong>vbUseDefault</strong> | -2                     | Use the setting from the computer's regional settings. |

 **Remarks**
When one or more optional arguments are omitted, the values for omitted arguments are provided by the computer's regional settings.

 **Note**  All settings information comes from the  **Regional Settings Number** tab.


