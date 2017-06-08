---
title: ViewCtl.SelectionChange Event (Outlook View Control)
ms.prod: outlook
ms.assetid: 4f637ff7-4b0d-c66e-ae51-bfd38b6e7f3a
ms.date: 06/08/2017
---


# ViewCtl.SelectionChange Event (Outlook View Control)

Occurs when the selection of the current view changes. 


## Syntax

 _expression_. **SelectionChange**

 _expression_A variable that represents a  **ViewCtl** object.


## Remarks

Other selection changes (such as the selected folder) do not cause this event to occur. 

This event does not occur if the current folder is a file system folder or if  **Outlook Today** or any folder with a current Web view is displayed.

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


