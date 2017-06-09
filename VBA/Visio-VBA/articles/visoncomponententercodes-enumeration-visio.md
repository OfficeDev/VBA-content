---
title: VisOnComponentEnterCodes Enumeration (Visio)
keywords: vis_sdr.chm70360
f1_keywords:
- vis_sdr.chm70360
ms.prod: visio
ms.assetid: ea4e67c3-58e5-6b6c-913b-58dec0a5448c
ms.date: 06/08/2017
---


# VisOnComponentEnterCodes Enumeration (Visio)

Flags to pass to the  **Application.OnComponentEnterState** method.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visComponentStateModal**|1|The state being identified is a modal state.|
| **visModalDeferEvents**|&;H10000|Causes Microsoft Visio to attempt to defer firing events while modal. By default, Visio defers firing events when displaying its own dialog boxes, but does not defer firing events when client code has caused a dialog box to appear.In the case of several nested modal scopes, if any scope is deferring events, all scopes within the outermost scope that is deferring events defer events.This flag only has an effect when Visio is entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.|
| **visModalDisableVisiosFrame**|&;H80000|Causes Visio to disable its frame window while modal. By default, Visio disables its frame window when showing its own dialog boxes or when showing dialog boxes implemented by Microsoft Visual Basic for Applications (VBA), but not when client code in another process shows a dialog box.If code in another process wants to show a dialog box and have the Visio frame window behave as if it is the Visio process showing the dialog box, it can set this flag.This flag only has an effect when entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.|
| **visModalDontBlockMessages**|&;H40000|Prevents Visio from rejecting calls from outside its main thread while modal. By default, Visio does reject calls from outside its thread while modal.In the case of several nested modal scopes, if any scope is deferring events, all scopes within the outermost scope that is deferring events defer events.This flag only has an effect when entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.|
| **visModalNoBeforeAfter**|&;H20000|Prevents Visio from firing a  **BeforeModal** event when entering a modal scope or an **AfterModal** event when leaving a modal scope.By default, Visio fires these events when displaying its own dialog boxes or displaying dialog boxes implemented by VBA, but does not fire these events when client code displays a dialog box.Calling the  **OnComponentEnterState** method causes these events to fire unless **visModalNoBeforeAfter** is specified.|

