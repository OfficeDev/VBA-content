---
title: Auto Macros
keywords: vbawd10.chm5209939
f1_keywords:
- vbawd10.chm5209939
ms.prod: word
ms.assetid: 65668dca-8517-5309-a89e-e19b3e85f4c6
ms.date: 06/08/2017
---


# Auto Macros

By giving a macro a special name, you can run it automatically when you perform an operation such as starting Word or opening a document. Word recognizes the following names as automatic macros, or "auto" macros.



|**Macro name**|**When it runs**|
|:-----|:-----|
|AutoExec|When you start Word or load a global template|
|AutoNew|Each time you create a new document|
|AutoOpen|Each time you open an existing document|
|AutoClose|Each time you close a document|
|AutoExit|When you exit Word or unload a global template|

Auto macros in code modules are recognized if either of the following conditions are true.


- The module is named after the auto macro (for example, AutoExec) and it contains a procedure named "Main."
    
- A procedure in any module is named after the auto macro.
    
Just like other macros, auto macros can be stored in the Normal template, another template, or a document. In order for an auto macro to run, it must be either in the Normal template, in the active document, or in the template on which the active document is based. The only exception is the AutoExec macro, which will not run automatically unless it is stored in one of the following: the Normal template, a template that is loaded globally through the  **Templates and Add-Ins** dialog box, or a global template stored in the folder specified as the Startup folder.
In the case of a naming conflict (multiple auto macros with the same name), Word runs the auto macro stored in the closest context. For example, if you create an AutoClose macro in a document and in the attached template, only the auto macro stored in the document will execute. If you create an AutoNew macro in the normal template, the macro will run if a macro named AutoNew does not exist in the document or the attached template.

 **Note**  You can hold down the SHIFT key to prevent auto macros from running. For example, if you create a new document based on a template that contains an AutoNew macro, you can prevent the AutoNew macro from running by holding down the SHIFT key when you click  **OK** in the **New** dialog box ( **File** menu) and continuing to hold down the SHIFT key until the new document is displayed. In a macro that might trigger an auto macro, you can use the following instruction to prevent auto macros from running.




```
WordBasic.DisableAutoMacros

```


