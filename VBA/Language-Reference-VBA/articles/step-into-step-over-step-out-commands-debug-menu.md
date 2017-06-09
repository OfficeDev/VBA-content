---
title: Step Into, Step Over, Step Out Commands (Debug Menu)
keywords: vbui6.chm2056453
f1_keywords:
- vbui6.chm2056453
ms.prod: office
ms.assetid: 5af5c030-adcc-f0a0-c4eb-33c3ba9a5789
ms.date: 06/08/2017
---


# Step Into, Step Over, Step Out Commands (Debug Menu)

 **Step Into**

Executes code one statement at a time.

When not in design mode,  **Step Into** enters[break mode](vbe-glossary.md) at the current line of execution. If the statement is a call to a procedure, the next statement displayed is the first statement in the procedure.

At [design time](vbe-glossary.md), this menu item begins execution and enters break mode before the first line of code is executed.
If there is no current execution point, the  **Step Into** command may appear to do nothing until you do something that triggers code, for example click on a document.
Toolbar button: 
![Toolbar button](images/tbr_stpi_ZA01201749.gif). Keyboard shortcut: F8.
 **Step Over**
Similar to  **Step Into**. The difference in use occurs when the current statement contains a call to a procedure.
 **Step Over** executes the procedure as a unit, and then steps to the next statement in the current procedure. Therefore, the next statement displayed is the next statement in the current procedure regardless of whether the current statement is a call to another procedure. Available in break mode only.
Toolbar button: 
![Toolbar button](images/tbr_stpo_ZA01201750.gif). Keyboard shortcut: SHIFT+F8.
 **Step Out**
Executes the remaining lines of a function in which the current execution point lies. The next statement displayed is the statement following the procedure call. All of the code is executed between the current and the final execution points. Available in break mode only.
Toolbar button: 
![Toolbar button](images/tbr_stot_ZA01201748.gif). Keyboard shortcut: CTRL+SHIFT+F8.

