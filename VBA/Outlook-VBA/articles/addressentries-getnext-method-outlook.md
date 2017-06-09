---
title: AddressEntries.GetNext Method (Outlook)
keywords: vbaol11.chm35
f1_keywords:
- vbaol11.chm35
ms.prod: outlook
api_name:
- Outlook.AddressEntries.GetNext
ms.assetid: 7579909c-90a2-660f-6cf5-039a441ccc93
ms.date: 06/08/2017
---


# AddressEntries.GetNext Method (Outlook)

Returns the next object in the  **[AddressEntries](addressentries-object-outlook.md)** collection.


## Syntax

 _expression_ . **GetNext**

 _expression_ A variable that represents an **AddressEntries** object.


### Return Value

An  **[AddressEntry](addressentry-object-outlook.md)** object that represents the next object contained by the collection.


## Remarks

It returns  **Nothing** if no next object exists, for example, if already positioned at the end of the collection.To ensure correct operation of the **[GetFirst](addressentries-getfirst-method-outlook.md)** , **[GetLast](addressentries-getlast-method-outlook.md)** , **GetNext** , and **[GetPrevious](addressentries-getprevious-method-outlook.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


[AddressEntries Object](addressentries-object-outlook.md)

