---
title: Application.System Property (Word)
keywords: vbawd10.chm158334985
f1_keywords:
- vbawd10.chm158334985
ms.prod: word
api_name:
- Word.Application.System
ms.assetid: 871f3821-4e17-1c63-9b4b-1d4e2bfc97d5
ms.date: 06/08/2017
---


# Application.System Property (Word)

Returns a  **[System](system-object-word.md)** object, which can be used to return system-related information and perform system-related tasks.


## Syntax

 _expression_ . **System**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example returns information about the system.


```
processor = System.ProcessorType 
enviro = System.OperatingSystem
```

This example establishes a connection to a network drive.




```
System.Connect Path:="\\Project\Info"
```


## See also


#### Concepts


[Application Object](application-object-word.md)

