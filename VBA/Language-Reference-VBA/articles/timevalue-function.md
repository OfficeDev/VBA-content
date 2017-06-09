---
title: TimeValue Function
keywords: vblr6.chm1009045
f1_keywords:
- vblr6.chm1009045
ms.prod: office
ms.assetid: 02ce264a-aa6b-2249-da37-dee3522c2db7
ms.date: 06/08/2017
---


# TimeValue Function



Returns a  **Variant** ( **Date** ) containing the time.
 **Syntax**
 **TimeValue(**_time_**)**
The required  _time_[argument](vbe-glossary.md) is normally a[string expression](vbe-glossary.md) representing a time from 0:00:00 (12:00:00 A.M.) to 23:59:59 (11:59:59 P.M.), inclusive. However, _time_ can also be any[expression](vbe-glossary.md) that represents a time in that range. If _time_ contains[Null](vbe-glossary.md),  **Null** is returned.
 **Remarks**
You can enter valid times using a 12-hour or 24-hour clock. For example,  `"2:24PM"` and `"14:24"` are both valid _time_ arguments.
If the  _time_ argument contains date information, **TimeValue** doesn't return it. However, if _time_ includes invalid date information, an error occurs.

## Example

This example uses the  **TimeValue** function to convert a string to a time. You can also use date literals to directly assign a time to a **Variant** or **Date** variable, for example, MyTime = #4:35:17 PM#.


```vb
Dim MyTime
MyTime = TimeValue("4:35:17 PM")    ' Return a time.

```


