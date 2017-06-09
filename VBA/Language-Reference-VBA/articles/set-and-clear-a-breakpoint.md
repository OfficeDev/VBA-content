---
title: Set and Clear a Breakpoint
keywords: vbhw6.chm1009020
f1_keywords:
- vbhw6.chm1009020
ms.prod: office
ms.assetid: 36b9640a-441a-0db8-aa03-5fda96215908
ms.date: 06/08/2017
---


# Set and Clear a Breakpoint

You set a [breakpoint](vbe-glossary.md) to suspend execution at a specific statement in a[procedure](vbe-glossary.md); for example, where you suspect problems may exist. You clear breakpoints when you no longer need them to stop execution.

 **To set a breakpoint**




1. Position the insertion point anywhere in a line of the [procedure](vbe-glossary.md) where you want execution to halt.
    
2. On the  **Debug** menu, click **Toggle Breakpoint** (F9), click next to the statement in the **Margin Indicator Bar** (if visible), or use the toolbar shortcut:
![Toolbar button](images/tbr_bkpt_ZA01201681.gif). The breakpoint is added and the line is set to the breakpoint color defined on the  **Editor Format** tab in the **Options** dialog box.
    

If you set a breakpoint on a line that contains several statements separated by colons ( **:** ), the break always occurs at the first statement on the line.
 **To clear a breakpoint**


1. Position the insertion point anywhere on a line of the procedure containing the breakpoint.
    
2. From the  **Debug** menu, choose **Toggle Breakpoint** (F9), or click next to the statement in the **Margin Indicator Bar** (if visible).
    
3. The breakpoint is cleared and highlighting is removed.
    

 **To clear all breakpoints in the application**


- From the  **Debug** menu, choose **Clear All Breakpoints** (CTRL+SHIFT+F9).
    
     **Note**  Breakpoints set in code are not saved when you save your code.


