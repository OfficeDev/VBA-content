---
title: Time Statement
keywords: vblr6.chm1009042
f1_keywords:
- vblr6.chm1009042
ms.prod: office
ms.assetid: 9c11edf2-5eac-207a-985e-1e990f3e1b12
ms.date: 06/08/2017
---


# Time Statement

Sets the system time.

 **Syntax**

 **Time =**_time_

The required  _time_[argument](vbe-glossary.md) is any[numeric expression](vbe-glossary.md), [string expression](vbe-glossary.md), or any combination, that can represent a time.
 **Remarks**
If  _time_ is a string, **Time** attempts to convert it to a time using the time separators you specified for your system. If it can't be converted to a valid time, an error occurs.

## Example

This example uses the  **Time** statement to set the computer system time to a user-defined time.


```vb
Dim MyTime 
MyTime = #4:35:17 PM# ' Assign a time. 
Time= MyTime ' Set system time to MyTime. 

```


