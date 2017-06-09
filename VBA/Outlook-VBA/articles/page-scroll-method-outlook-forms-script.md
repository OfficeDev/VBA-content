---
title: Page.Scroll Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 19810ddc-70f9-aa60-0c8a-f2c9ab9ff51f
ms.date: 06/08/2017
---


# Page.Scroll Method (Outlook Forms Script)

Moves the scroll bar on an object.


## Syntax

 _expression_. **Scroll**( **_xAction_**,  **_yAction_**)

 _expression_A variable that represents a  **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|xAction|Optional| **Variant**|Identifies the action to occur in the horizontal direction.|
|yAction|Optional| **Variant**|Identifies the action to occur in the vertical direction.|

## Remarks

The settings for  _xAction_ and _yAction_ are:



|**Value**|**Description**|
|:-----|:-----|
|0|Do not scroll in the specified direction.|
|1|Move up on a vertical scroll bar or left on a horizontal scroll bar. Movement is equivalent to pressing the up or left arrow key on the keyboard to move the scroll bar.|
|2|Move down on a vertical scroll bar or right on a horizontal scroll bar. Movement is equivalent to pressing the right or down arrow key on the keyboard to move the scroll bar.|
|3|Move one pageup on a vertical scroll bar or one page left on a horizontal scroll bar. Movement is equivalent to pressing  **PAGE UP** on the keyboard to move the scroll bar.|
|4|Move one pagedown on a vertical scroll bar or one page right on a horizontal scroll bar. Movement is equivalent to pressing  **PAGE DOWN** on the keyboard to move the scroll bar.|
|5|Move to the top of a vertical scroll bar or to the left end of a horizontal scroll bar.|
|6|Move to the bottom of a vertical scroll bar or to the right end of a horizontal scroll bar.|
The  **Scroll** method applies scroll bars that appear on a **[Page](page-object-outlook-forms-script.md)** that is larger than its display area. This method does not apply to the stand-alone **[ScrollBar](scrollbar-object-outlook-forms-script.md)** control or to scroll bars that appear on a **[TextBox](textbox-object-outlook-forms-script.md)**.


