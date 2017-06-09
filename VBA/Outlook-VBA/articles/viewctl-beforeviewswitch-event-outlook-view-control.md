---
title: ViewCtl.BeforeViewSwitch Event (Outlook View Control)
ms.prod: outlook
ms.assetid: f68c1cd3-7463-0e2b-7fee-d5a100b79f8c
ms.date: 06/08/2017
---


# ViewCtl.BeforeViewSwitch Event (Outlook View Control)

Occurs before Microsoft Outlook changes the view that is applied to the folder displayed in the View Control element, either as a result of user action or through program code. 


## Syntax

 _expression_. **BeforeViewSwitch**( **_newView_**,  **_Cancel_**)

 _expression_A variable that represents a  **ViewCtl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|newView|Required| **String**|The name of the view that the View Control is switching to.|
|Cancel|Optional| **Boolean**| **False** when the event occurs. If the event procedure sets this parameter to **True**, it cancels the switch and retains the current view.|

## Remarks

You can cancel this event to prevent the user from changing the view in the View Control. 

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


