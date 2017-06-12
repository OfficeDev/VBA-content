---
title: CreateEventProc Method (VBA Add-In Object Model)
keywords: vbob6.chm104021
f1_keywords:
- vbob6.chm104021
ms.prod: office
ms.assetid: afcdc0a2-aa3d-6882-f89c-17f0dcf3df2b
ms.date: 06/08/2017
---


# CreateEventProc Method (VBA Add-In Object Model)



Creates an event [procedure](vbe-glossary.md).
 **Syntax**
 _object_**.CreateEventProc(**_eventname_, _objectname_**) As Long**
The  **CreateEventProc** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _eventname_|Required. A [string expression](vbe-glossary.md) specifying the name of the event you want to add to the[module](vbe-glossary.md).|
| _objectname_|Required. A string expression specifying the name of the object that is the source of the event.|
 **Remarks**
Use the  **CreateEventProc** method to create an event procedure. For example, to create an event procedure for the **Click** event of a **Command Button** control named `Command1` you would use the following code, where `CM` represents an object of type **CodeModule**:



```
TextLocation = CM.CreateEventProc("Click", "Command1")
```

The  **CreateEventProc** method returns the line at which the body of the event procedure starts. **CreateEventProc** fails if the[arguments](vbe-glossary.md) refer to a nonexistent event.

