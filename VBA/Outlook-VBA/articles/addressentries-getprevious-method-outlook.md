---
title: AddressEntries.GetPrevious Method (Outlook)
keywords: vbaol11.chm36
f1_keywords:
- vbaol11.chm36
ms.prod: outlook
api_name:
- Outlook.AddressEntries.GetPrevious
ms.assetid: 3d5aa211-212e-9a97-58aa-47d4447c9f47
ms.date: 06/08/2017
---


# AddressEntries.GetPrevious Method (Outlook)

Returns the previous object in the  **[AddressEntries](addressentries-object-outlook.md)** collection.


## Syntax

 _expression_ . **GetPrevious**

 _expression_ A variable that represents an **AddressEntries** object.


### Return Value

An  **[AddressEntry](addressentry-object-outlook.md)** object that represents the previous object contained by the collection.


## Remarks

It returns  **Nothing** if no previous object exists, for example, if already positioned at the beginning of the collection.To ensure correct operation of the **[GetFirst](addressentries-getfirst-method-outlook.md)** , **[GetLast](addressentries-getlast-method-outlook.md)** , **[GetNext](addressentries-getnext-method-outlook.md)** , and **GetPrevious** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


[AddressEntries Object](addressentries-object-outlook.md)

