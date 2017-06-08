---
title: Field Events
keywords: olfm10.chm3077125
f1_keywords:
- olfm10.chm3077125
ms.prod: outlook
ms.assetid: 05b13be0-c964-26a7-995a-7a74629026f3
ms.date: 06/08/2017
---


# Field Events



Outlook provides two events to notify your program that a field (property) in an item has changed. The  **PropertyChange** event is fired whenever a standard Outlook field in an item has changed. Outlook fires the **CustomPropertyChange** event whenever a user-defined field changes.
A control that is bound to a field does not fire the  **Click** event, whether the control was selected from the **Control Toolbox** and subsequently bound to a field, or was selected from the **Field Chooser**. Consequently, you must use the  **PropertyChange** or **CustomPropertyChange** event to detect user interaction with a bound control.

