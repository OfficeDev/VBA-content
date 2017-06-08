---
title: Application.Help Method (Word)
keywords: vbawd10.chm158335305
f1_keywords:
- vbawd10.chm158335305
ms.prod: word
api_name:
- Word.Application.Help
ms.assetid: ff64e6bd-e29b-7cfc-437b-df8b8e59ce59
ms.date: 06/08/2017
---


# Application.Help Method (Word)

Displays installed Help information.


## Syntax

 _expression_ . **Help**( **_HelpType_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HelpType_|Required| **Variant**|The on-line Help topic or window. Can be any of these  **[WdHelpType](wdhelptype-enumeration-word.md)** constants: **wdHelp** , **wdHelpAbout** , **wdHelpActiveWindow** , **wdHelpContents** , **wdHelpHWP** , **wdHelpIchitaro** , **wdHelpIndex** , **wdHelpPE2** , **wdHelpPSSHelp** , **wdHelpSearch** , **wdHelpUsingHelp** . (Some of the constants listed here may not be available to you, depending on the language that you have selected or installed.)|

## Example

This example displays the  **Help Topics** dialog box.


```
Help HelpType:=wdHelp
```

This example displays a list of Help topics that describe how to use Help.




```
Help HelpType:=wdHelpUsingHelp
```


## See also


#### Concepts


[Application Object](application-object-word.md)

