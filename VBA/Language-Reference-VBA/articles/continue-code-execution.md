---
title: Continue Code Execution
keywords: vbhw6.chm1008878
f1_keywords:
- vbhw6.chm1008878
ms.prod: office
ms.assetid: 61035245-f12f-dea4-fa8e-5904f34d1bf3
ms.date: 06/08/2017
---


# Continue Code Execution

When you run your code, execution may stop if:



- An untrapped [run-time error](vbe-glossary.md) occurs.
    
- A trapped run-time error occurs, and  **Break on All Errors** is selected on the **General** tab of the **Options** dialog box ( **Tools** menu).
    
- A previously set [breakpoint](vbe-glossary.md) is encountered.
    
- A  **Stop** statement in your code is encountered, switching the mode to[break mode](vbe-glossary.md).
    
- An  **End** statement in your code is encountered, switching the mode to[design time](vbe-glossary.md).
    
- You halt execution manually at a given point.
    
- A [watch expression](vbe-glossary.md), which you set to break when the value has changed or break when the value is true, is encountered.
    

 **To halt execution manually**


1. To switch to break mode, choose  **Break** (CTRL+BREAK) from the **Run** menu, or use the toolbar shortcut:
![Toolbar button](images/tbr_brk_ZA01201682.gif).
    
2. To switch to design time, choose  **Reset <projectname&gt;** from the **Run** menu, or use the toolbar shortcut:
![Toolbar button](images/tbr_end_ZA01201701.gif).
    

 **To continue execution when your application has halted**


- On the  **Run** menu, click **Continue** (F5), or use the toolbar shortcut:
![Toolbar button](images/tbr_strt_ZA01201751.gif). - Or -
    
- On the  **Debug** menu, click **Step Into** (F8), **Step Over** (SHIFT+F8), **Step Out** (CTRL+SHIFT+F8), or **Run To Cursor** (CTRL+F8)(.
    

 **To continue execution when your application has halted because of a handled error**


- Press ALT+F8 to step through the error-handler. - Or -
    
- Press ALT+F5 to resume execution by running through the error-handler.
    


