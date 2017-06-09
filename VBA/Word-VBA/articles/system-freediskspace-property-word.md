---
title: System.FreeDiskSpace Property (Word)
keywords: vbawd10.chm154468356
f1_keywords:
- vbawd10.chm154468356
ms.prod: word
api_name:
- Word.System.FreeDiskSpace
ms.assetid: 739db138-37f3-821b-8214-013153b20fa0
ms.date: 06/08/2017
---


# System.FreeDiskSpace Property (Word)

Returns the available disk space for the current drive, in bytes. Use the ChDrive statement to change the current drive. Read-only  **Long** .


## Syntax

 _expression_ . **FreeDiskSpace**

 _expression_ A variable that represents a **[System](system-object-word.md)** object.


## Remarks

There are 1024 bytes in a kilobyte and 1,048,576 bytes in a megabyte. The maximum return value for the  **FreeDiskSpace** property is 2,147,483,647. Therefore, even if you have four gigabytes of free disk space, it returns 2,147,483,647.


## Example

This example checks the amount of free disk space. If there is less than 10 megabytes of space available, a message is displayed.


```vb
If (System.FreeDiskSpace \ 1048576) < 10 Then _ 
 MsgBox "Low disk space"
```


## See also


#### Concepts


[System Object](system-object-word.md)

