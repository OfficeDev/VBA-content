---
title: Call Stack Command (View Menu)
keywords: vbui6.chm2007505
f1_keywords:
- vbui6.chm2007505
ms.prod: office
ms.assetid: c5eaa876-9fbc-370c-ec17-9c04b7fc6854
ms.date: 06/08/2017
---


# Call Stack Command (View Menu)

Displays the  **Call Stack** dialog box, which lists the procedures that have started but are not completed. Available only in[break mode](vbe-glossary.md).

When Visual Basic is executing the code in a procedure, that procedure is added to a list of active procedure calls. If that procedure then calls another procedure, there are two procedures on the list of active procedure calls. Each time a procedure calls another [Sub](vbe-glossary.md), [Function](vbe-glossary.md), or  **Property** procedure, it is added to the list. Each procedure is removed from the list as execution is returned to the calling procedure. Procedures called from the **Immediate** window are also added to the calls list.

You can also display the  **Call Stack** dialog box by clicking the **Calls** button (...) next to the Procedure box in the **Locals** window.

Toolbar shortcut: 
![Toolbar button](images/tbr_call_ZA01201683.gif). Keyboard shortcut: CTRL+L.

