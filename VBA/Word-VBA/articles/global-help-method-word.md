---
title: Global.Help Method (Word)
keywords: vbawd10.chm163119433
f1_keywords:
- vbawd10.chm163119433
ms.prod: word
api_name:
- Word.Global.Help
ms.assetid: cfae6e61-84bf-2462-39c5-569baec866ee
ms.date: 06/08/2017
---


# Global.Help Method (Word)

Displays on-line Help information.


## Syntax

 _expression_ . **Help**( **_HelpType_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HelpType_|Required| **Variant**|The on-line Help topic or window. Can be any of these  **[WdHelpType](wdhelptype-enumeration-word.md)** constants.|

## Remarks

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.


## Example

This example displays the Help Topics dialog box.


```
Help HelpType:=wdHelp
```

This example displays a list of Help topics that describe how to use Help.




```
Help HelpType:=wdHelpUsingHelp
```


## See also


#### Concepts


[Global Object](global-object-word.md)

