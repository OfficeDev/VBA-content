---
title: System.Version Property (Word)
keywords: vbawd10.chm154468355
f1_keywords:
- vbawd10.chm154468355
ms.prod: word
api_name:
- Word.System.Version
ms.assetid: 0d937656-94eb-2fa5-0d00-bfdfeae59ecf
ms.date: 06/08/2017
---


# System.Version Property (Word)

Returns the version number of the operating system. Read-only  **String** .


## Syntax

 _expression_ . **Version**

 _expression_ A variable that represents a **[System](system-object-word.md)** object.


## Example

This example displays the version number of the operating system in a message box.


```
Msgbox "The system version is " &; System.Version
```


## See also


#### Concepts


[System Object](system-object-word.md)

