---
title: Conflicts.GetNext Method (Outlook)
keywords: vbaol11.chm408
f1_keywords:
- vbaol11.chm408
ms.prod: outlook
api_name:
- Outlook.Conflicts.GetNext
ms.assetid: 2e21ea88-c732-17ee-cd87-698fee992269
ms.date: 06/08/2017
---


# Conflicts.GetNext Method (Outlook)

Returns the next object in the  **[Conflicts](conflicts-object-outlook.md)** collection.


## Syntax

 _expression_ . **GetNext**

 _expression_ A variable that represents a **Conflicts** object.


### Return Value

A  **[Conflict](conflict-object-outlook.md)** object that represents the next object contained by the collection.


## Remarks

It returns  **Nothing** if no next object exists, for example, if already positioned at the end of the collection. To ensure correct operation of the **[GetFirst](conflicts-getfirst-method-outlook.md)** , **[GetLast](conflicts-getlast-method-outlook.md)** , **GetNext** , and **[GetPrevious](conflicts-getprevious-method-outlook.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


[Conflicts Object](conflicts-object-outlook.md)

