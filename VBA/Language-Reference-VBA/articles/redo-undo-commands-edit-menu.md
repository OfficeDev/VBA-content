---
title: Redo, Undo Commands (Edit Menu)
keywords: vbui6.chm2012587
f1_keywords:
- vbui6.chm2012587
ms.prod: office
ms.assetid: 63a4524d-62b9-d97c-e79b-1b38dcdfd073
ms.date: 06/08/2017
---


# Redo, Undo Commands (Edit Menu)

 **Undo**

Reverses the last editing action, such as typing text in the  **Code** window or deleting controls. When you delete one or more controls, you can use the **Undo** command to restore the controls and all their properties.

Toolbar shortcut 
![Toolbar button](images/tbr_undo_ZA01201762.gif). Keyboard shortcuts: CTRL+Z or ALT+BACKSPACE.


 **Note**  You can't undo a  **Cut** operation using the **Undo** command on a form.

 **Redo**
Restores the last text editing or resizing and positioning of controls if no other actions have occurred since the last  **Undo**.
Toolbar shortcut 
![Toolbar button](images/tbr_redo_ZA01201734.gif).
For text edits, you can use  **Undo** and **Redo** to restore up to twenty edits.
These commands are unavailable at runtime, or if there was no previous edit, or if any other action has been performed after the last edit. Also, some large edits may cause low memory conditions that could prevent an  **Undo** action.

