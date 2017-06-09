---
title: SpellingOptions.HebrewModes Property (Excel)
keywords: vbaxl10.chm717083
f1_keywords:
- vbaxl10.chm717083
ms.prod: excel
api_name:
- Excel.SpellingOptions.HebrewModes
ms.assetid: b8ecfa29-7ec4-180b-fb37-6876ab6c0cc7
ms.date: 06/08/2017
---


# SpellingOptions.HebrewModes Property (Excel)

Returns or sets the mode for the Hebrew spelling checker. Read/write  **[XlHebrewModes](xlhebrewmodes-enumeration-excel.md)** .


## Syntax

 _expression_ . **HebrewModes**

 _expression_ A variable that represents a **SpellingOptions** object.


## Remarks



| **XlHebrewModes** can be one of these **XlHebrewModes** constants.|
| **xlHebrewFullScript** (default) The conventional script type as required by the Hebrew Language Academy when writing non-diacritisized text.|
| **xlHebrewMixedAuthorizedScript** The Hebrew traditional script.|
| **xlHebrewMixedScript** In this mode the speller accepts any word recognized as Hebrew, whether in Full Script, Partial Script, or any non-conventional spelling variation what is known to the speller.|
| **xlHebrewPartialScript** In this mode the speller accepts words both in Full Script and Partial Script. Some words will be flagged since this spelling is not authorized in either Full script or Partial script.|
A legitimate Hebrew word can be a basic dictionary entry or any inflection.


## Example

In this example, Microsoft Excel determines the setting for the Hebrew spelling mode and notifies the user.


```vb
Sub CheckHebrewMode() 
 
 ' Determine the Hebrew spelling mode setting and notify user. 
 Select Case Application.SpellingOptions.HebrewModes 
 Case xlHebrewFullScript 
 MsgBox "The Hebrew spelling mode setting is Full Script." 
 Case xlHebrewMixedAuthorizedScript 
 MsgBox "The Hebrew spelling mode setting is Mixed Authorized Script." 
 Case xlHebrewMixedScript 
 MsgBox "The Hebrew spelling mode setting is Mixed Script." 
 Case xlHebrewPartialScript 
 MsgBox "The Hebrew spelling mode setting is Partial Script." 
 End Select 
 
End Sub
```


## See also


#### Concepts


[SpellingOptions Object](spellingoptions-object-excel.md)

