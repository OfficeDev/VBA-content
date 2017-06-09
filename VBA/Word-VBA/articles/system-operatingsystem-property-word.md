---
title: System.OperatingSystem Property (Word)
keywords: vbawd10.chm154468353
f1_keywords:
- vbawd10.chm154468353
ms.prod: word
api_name:
- Word.System.OperatingSystem
ms.assetid: 471183cf-ac38-c6ab-c468-05ed35b10b9b
ms.date: 06/08/2017
---


# System.OperatingSystem Property (Word)

Returns the name of the current operating system (for example, "Windows" or "Windows NT"). Read-only  **String** .


## Syntax

 _expression_ . **OperatingSystem**

 _expression_ An expression that returns a **[System](system-object-word.md)** object.


## Example

This example displays a message that includes the name of the current operating system.


```vb
MsgBox "This computer is running " &; System.OperatingSystem
```


## See also


#### Concepts


[System Object](system-object-word.md)

