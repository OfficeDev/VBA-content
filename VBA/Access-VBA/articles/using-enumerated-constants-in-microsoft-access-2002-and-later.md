---
title: Using Enumerated Constants in Microsoft Access 2002 and Later
keywords: vbaac10.chm113266
f1_keywords:
- vbaac10.chm113266
ms.prod: access
ms.assetid: 7eb8fb08-76e5-a59f-5d6d-64c7081470e6
ms.date: 06/08/2017
---


# Using Enumerated Constants in Microsoft Access 2002 and Later

In Access 2002, a number of intrinsic constants were added or changed. This was done to create lists of "enumerated" constants that are displayed in the  **Auto List Members** list in the Module window for the arguments of various Access methods, functions, and properties, or as the setting of various Access properties. You can select the appropriate constant from the list in the Module window, instead of having to remember the constant or look it up in the Help topic.

The following information applies to enumerated constants:

- The set of enumerated constants for each method, function, or property argument has a name, which is displayed in the syntax line for the method, function, or property in the Module window when the  **Auto Quick Info** option is selected in the **Editor** tab of the **Options** dialog box, available by clicking **Options** on the **Tools** menu. (For property settings, the name isn't displayed, just the list of constants.) For example, the syntax line for the **[OpenForm](docmd-openform-method-access.md)** method of the **[DoCmd](docmd-object-access.md)** object shows **[View As AcFormView = acNormal]** for the _view_ argument of this method. **AcFormView** is the name of this set of enumerated constants, and **acNormal** is the default setting for the argument. The Object Browser also lists the names of the sets of enumerated constants in the **Classes** box and lists the intrinsic constants contained in each of these sets in the **Members Of** box.
    
- For constant names that have changed, the old constants will still work. For example, one of the intrinsic constants for the  _save_ argument of the **Close** method of the **DoCmd** object was **acPrompt**. It's now **acSavePrompt**, but **acPrompt** will still work.
    
- In a number of cases in previous versions of Access, you could leave an argument setting blank, and Access would perform the default action for that argument. For example, you could leave the  _objecttype_ (and _objectname_ ) arguments of the **Close** method blank, and Access would close the active window. For the new sets of enumerated constants, the blank setting has been replaced with a new default constant. For example, the _objecttype_ argument of the **Close** method now has a new default constant, **acDefault**. Setting this argument to the new constant has the same effect as leaving the argument blank. In addition, you can still leave such arguments blank, and Access will assume the new default constant.
    
- There's one exception to this. If you run Visual Basic code from previous versions of Visual Basic in Access by using Automation, blank arguments will cause an error for those arguments that have the new default constants. This problem doesn't occur for old Visual Basic for Applications or Visual Basic code run directly in Access.
    

