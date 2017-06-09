---
title: Stop Code Execution
keywords: vbhw6.chm1008937
f1_keywords:
- vbhw6.chm1008937
ms.prod: office
ms.assetid: f0608ca1-d6d8-d722-cbd7-8a31634264ed
ms.date: 06/08/2017
---


# Stop Code Execution

As you run your code, it may stop executing for one of the following reasons:



- An untrapped [run-time error](vbe-glossary.md) occurs.
    
- A trapped run-time error occurs, and  **Break on All Errors** is selected on the **General** tab in the **Options** dialog box.
    
- A [breakpoint](vbe-glossary.md) is encountered.
    
- A  **Stop** statement is encountered in your code, switching the mode to[break mode](vbe-glossary.md).
    
- An  **End** statement is encountered in your code, switching the mode to[design time](vbe-glossary.md).
    
- You halt execution manually at a given point.
    
- A [watch expression](vbe-glossary.md) that you set to break if its value changes or becomes true is encountered.
    

 **To halt execution manually**


- To switch to break mode, from the  **Run** menu, choose **Break** (CTRL+BREAK), or use the toolbar shortcut:
![Toolbar button](images/tbr_brk_ZA01201682.gif).
    
- To switch to design time, from the  **Run** menu, choose **Reset <projectname&gt;**, or use the toolbar shortcut:
![Toolbar button](images/tbr_end_ZA01201701.gif).
    

 **To continue execution when your application has halted**


- From the  **Debug** menu, choose **Step Into** (F8), **Step Over** (SHIFT+F8), **Step Out** (CTRL+SHIFT+F8), or **Run To Cursor** (CTRL+F8.
    


