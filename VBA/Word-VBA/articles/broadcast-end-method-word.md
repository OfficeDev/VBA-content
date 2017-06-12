---
title: Broadcast.End Method (Word)
keywords: vbawd10.chm36438120
f1_keywords:
- vbawd10.chm36438120
ms.prod: word
ms.assetid: dca52c1c-c337-f9ee-6c82-ef05da5cdf45
ms.date: 06/08/2017
---


# Broadcast.End Method (Word)

Ends the specified broadcast session.


## Syntax

 _expression_ . **End**

 _expression_ A variable that represents a **Broadcast** object.


### Return value

 **VOID**


## Remarks

Calling the  **End** method terminates the broadcast session without displaying a confirmation prompt to the user. It also sets the value of the[Broadcast.AttendeeURL](broadcast-attendeeurl-property-word.md) property to an empty string.

If the document is not being broadcast, the method returns runtime error 4702.


## See also


#### Other resources


[Broadcast Object](broadcast-object-word.md)


