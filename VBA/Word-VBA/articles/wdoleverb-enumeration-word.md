---
title: WdOLEVerb Enumeration (Word)
ms.prod: word
api_name:
- Word.WdOLEVerb
ms.assetid: 0a5ef4a2-0982-8fb8-7173-39286c599e6a
ms.date: 06/08/2017
---


# WdOLEVerb Enumeration (Word)

Specifies the action associated with the verb that the OLE object should perform.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdOLEVerbDiscardUndoState**|-6|Forces the object to discard any undo state that it might be maintaining; note that the object remains active, however.|
| **wdOLEVerbHide**|-3|Removes the object's user interface from view.|
| **wdOLEVerbInPlaceActivate**|-5|Runs the object and installs its window, but doesn't install any user-interface tools.|
| **wdOLEVerbOpen**|-2|Opens the object in a separate window.|
| **wdOLEVerbPrimary**|0|Performs the verb that is invoked when the user double-clicks the object.|
| **wdOLEVerbShow**|-1|Shows the object to the user for editing or viewing. Use it to show a newly inserted object for initial editing.|
| **wdOLEVerbUIActivate**|-4|Activates the object in place and displays any user-interface tools that the object needs, such as menus or toolbars.|

