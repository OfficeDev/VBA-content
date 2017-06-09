---
title: TabStrip.Click Event (Outlook Forms Script)
ms.prod: outlook
ms.assetid: d79676f8-eb45-1fc0-e631-4f7f79e4f418
ms.date: 06/08/2017
---


# TabStrip.Click Event (Outlook Forms Script)

Occurs when the user clicks inside the control.


## Syntax

 _expression_. **Click**( **_Index_**, )

 _expression_A variable that represents a  **TabStrip** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**||

## Remarks

The following are examples of actions that initiate the  **Click** event of the specified control:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a control with the left mouse button (left-clicking).
    
- Pressing a control's accelerator key.
    


The  **Click** event is not initiated when **Value** is set to **Null**.

Left-clicking changes the value of a control, thus it initiates the  **Click** event. Right-clicking does not change the value of the control, so it does not initiate the **Click** event.


