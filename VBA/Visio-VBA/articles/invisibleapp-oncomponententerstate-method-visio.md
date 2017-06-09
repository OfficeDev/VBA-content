---
title: InvisibleApp.OnComponentEnterState Method (Visio)
keywords: vis_sdr.chm17552045
f1_keywords:
- vis_sdr.chm17552045
ms.prod: visio
api_name:
- Visio.InvisibleApp.OnComponentEnterState
ms.assetid: 4550b7cf-3aaa-cfba-edf0-662847d7e970
ms.date: 06/08/2017
---


# InvisibleApp.OnComponentEnterState Method (Visio)

Informs a Microsoft Visio instance that client code is causing the instance to enter or exit a particular state.


## Syntax

 _expression_ . **OnComponentEnterState**( **_uStateID_** , **_bEnter_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _uStateID_|Required| **VisOnComponentEnterCodes**|Describes the state being entered or exited along with flags that influence behavior while in the indicated state.|
| _bEnter_|Required| **Boolean**| **True** to indicate that a state is being entered; **False** to indicate that a state is being exited.|

### Return Value

nothing


## Remarks

The  _uStateID_argument indicates the state being entered or exited. Code that calls this method should do so both when entering and exiting the state.

At present, the only state change supported by the  **OnComponentEnterState** method is **visComponentStateModal** , indicating that the client is performing an action that causes Visio to enter or exit a modal state.

Most client code doesn't need to call the  **OnComponentEnterState** method when causing Visio to enter or leave the state of being modal, for example, when showing modal dialog boxes. Typically, this method is used by client code that shows dialog boxes other than Microsoft Visual Basic for Applications (VBA) forms and requires behavior different from the Visio default behavior.

Following are constants and values for  _uStateID_, which are declared by the Visio type library in  **VisOnComponentEnterCodes** . Any of the following constants prefixed with **visModal** can be combined with **visComponentStateModal** to influence Visio behavior when transitioning to or from a modal state.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visComponentStateModal**|1|The state being identified is a modal state.|
| **visModalDeferEvents**|&;H10000|Causes Visio to attempt to defer firing events while modal. By default, Visio defers firing events when displaying its own dialog boxes, but does not defer firing events when client code has caused a dialog box to appear.In the case of several nested modal scopes, if any scope is deferring events, all scopes within the outermost scope that is deferring events will defer events.This flag only has an effect when Visio is entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.|
| **visModalNoBeforeAfter**|&;H20000|Prevents Visio from firing a  **BeforeModal** event when entering a modal scope or an **AfterModal** event when leaving a modal scope.By default, Visio fires these events when displaying its own dialog boxes or displaying dialog boxes implemented by VBA, but does not fire these events when client code displays a dialog box.Calling the **OnComponentEnterState** method will cause these events to fire unless **visModalNoBeforeAfter** is specified.|
| **visModalDontBlockMessages**|&;H40000|Prevents Visio from rejecting calls from outside its main thread while modal. By default, Visio does reject calls from outside its thread while modal.In the case of several nested modal scopes, if any scope is deferring events, all scopes within the outermost scope that is deferring events will defer events.This flag only has an effect when Visio is entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.|
| **visModalDisableVisiosFrame**|&;H80000|Causes Visio to disable its frame window while modal. By default, Visio disables its frame window when showing its own dialog boxes or when showing dialog boxes implemented by VBA, but not when client code in another process shows a dialog box.If code in another process wants to show a dialog box and have the Visio frame window behave as if it is the Visio process showing the dialog box, it can set this flag.This flag only has an effect when Visio entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.|

