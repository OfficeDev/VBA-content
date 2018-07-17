---
title: Avoiding Naming Conflicts
keywords: vbcn6.chm1076671
f1_keywords:
- vbcn6.chm1076671
ms.prod: office
ms.assetid: bdeffd46-cdbc-4702-472a-e28df9507bb1
ms.date: 06/08/2017
---


# Avoiding Naming Conflicts

A naming conflict occurs when you try to create or use an [identifier](vbe-glossary.md) that was previously defined. In some cases, naming conflicts generate errors such as " `Ambiguous name detected`" or "" or " `Duplicate declaration in current scope`". Naming conflicts that go undetected can result in bugs in your code that produce erroneous results, especially if you do not explicitly declare all [variables](vbe-glossary.md) before first use.

You can avoid most naming conflicts by understanding the [scoping](vbe-glossary.md), private [module-level](vbe-glossary.md), and public module-level.

A naming conflict can occur when an identifier:


- Is visible at more than one scoping level.
    
- Has two different meanings at the same level.
    

For example, procedures in separate [modules](vbe-glossary.md) can have the same name. Therefore, you can define a procedure named `MySub` in modules named `Mod1` and `Mod2`. No conflicts occur if each procedure is called only from other procedures in its own module. However, an error can occur if  `MySub` is called from a third module, and no qualification is provided to distinguish between the two `MySub` procedures.
Most naming conflicts can be resolved by preceding each identifier with a qualifier that consists of the module name and, if necessary, a [project](vbe-glossary.md) name. For example:



```
YourProject.YourModule.YourSub MyProject.MyModule.MyVar
```

The preceding code calls the  **Sub** procedure `YourSub` and passes `MyVar` as an argument. You can use any combination of qualifiers to differentiate identical identifiers.
Visual Basic matches each reference to an identifier with the "closest" declaration of a matching identifier. For example, if  `MyID` is declared **Public** in two modules in a project ( `Mod1` and `Mod2`), you can specify the  `MyID` declared in `Mod2` without qualification from within `Mod2`, but you must qualify it as  `Mod2.MyID` to specify it in `Mod1`. This is also true if  `Mod2` is in a different but directly [referenced project](vbe-glossary.md). However, if  `Mod2` is in an indirectly referenced project, that is, a project referenced by the project you directly reference, references to the `Mod2` variable named `MyID` must always be qualified with the project name. If you reference `MyID` from a third, directly referenced module, the match is made with the first declaration encountered by searching:

- Directly referenced projects, in the order that they appear in the  **References** dialog box of the **Tools** menu.
    
- The modules of each project. Note that there is no inherent order to the modules in the project.
    
You can't reuse names of [host-application](vbe-glossary.md) objects, for example, R1C1 in Microsoft Excel, at different scoping levels.

 **Tip**  Typical errors caused by naming conflicts include ambiguous names, duplicate declarations, undeclared identifiers, and procedures that are not found. By beginning each module with an  **Option Explicit** statement to force explicit declarations of variables before they are used, you can avoid some potential naming conflicts and identifier-related bugs.


