---
title: Explorer.SelectAllItems Method (Outlook)
keywords: vbaol11.chm3311
f1_keywords:
- vbaol11.chm3311
ms.prod: outlook
api_name:
- Outlook.Explorer.SelectAllItems
ms.assetid: 05b3169a-5f27-2169-5ac5-1d64951d6430
ms.date: 06/08/2017
---


# Explorer.SelectAllItems Method (Outlook)

Selects all items that are contained in the current view of the active explorer. 


## Syntax

 _expression_ . **SelectAllItems**

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


## Remarks

If one or more groups are collapsed in the current view, calling  **SelectAllItems** does not select items in the collapsed groups. Only items in expanded groups are selected.

If the current view is a calendar view, calling  **SelectAllItems** selects all appointments and all-day events in the view. Calling **SelectAllItems** on a calendar view does not select items in the Daily Task List.

The  **[SelectionChange](explorer-selectionchange-event-outlook.md)** event fires only once after the **SelectAllItems** method is called.

If the current view or current folder does not contain any items, calling  **SelectAllItems** does not result in any change to the selection and does not fire the **SelectionChange** event.

 **SelectAllItems** raises an error if the item is being edited in the current view, or the current view is a conversation view.


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

