---
title: IMEStatus Function
keywords: vblr6.chm1011064
f1_keywords:
- vblr6.chm1011064
ms.prod: office
ms.assetid: 8fb525b1-0243-79c8-32fc-4eb8d634e351
ms.date: 06/08/2017
---


# IMEStatus Function



Returns an [Integer](vbe-glossary.md) specifying the current Input Method Editor (IME) mode of Microsoft Windows; available in East Asian versions only.
 **Syntax**
 **IMEStatus**
 **Return Values**
The return values for the Japanese [locale](vbe-glossary.md) are as follows:


| <strong>Constant</strong>              | <strong>Value</strong> | <strong>Description</strong> |
|:---------------------------------------|:-----------------------|:-----------------------------|
| <strong>vbIMEModeNoControl</strong>    | 0                      | Don't control IME (default)  |
| <strong>vbIMEModeOn</strong>           | 1                      | IME on                       |
| <strong>vbIMEModeOff</strong>          | 2                      | IME off                      |
| <strong>vbIMEModeDisable</strong>      | 3                      | IME disabled                 |
| <strong>vbIMEModeHiragana</strong>     | 4                      | Full-width Hiragana mode     |
| <strong>vbIMEModeKatakana</strong>     | 5                      | Full-width Katakana mode     |
| <strong>vbIMEModeKatakanaHalf</strong> | 6                      | Half-width Katakana mode     |
| <strong>vbIMEModeAlphaFull</strong>    | 7                      | Full-width Alphanumeric mode |
| <strong>vbIMEModeAlpha</strong>        | 8                      | Half-width Alphanumeric mode |

The return values for the Korean locale are as follows:


| <strong>Constant</strong>            | <strong>Value</strong> | <strong>Description</strong> |
|:-------------------------------------|:-----------------------|:-----------------------------|
| <strong>vbIMEModeNoControl</strong>  | 0                      | Don't control IME(default)   |
| <strong>vbIMEModeAlphaFull</strong>  | 7                      | Full-width Alphanumeric mode |
| <strong>vbIMEModeAlpha</strong>      | 8                      | Half-width Alphanumeric mode |
| <strong>vbIMEModeHangulFull</strong> | 9                      | Full-width Hangul mode       |
| <strong>vbIMEModeHangul</strong>     | 10                     | Half-width Hangul mode       |

The return values for the Chinese locale are as follows:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbIMEModeNoControl**|0|Don't control IME (default)|
|**vbIMEModeOn**|1|IME on|
|**vbIMEModeOff**|2|IME off|

