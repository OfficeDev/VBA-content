---
title: Improvements in Compilation Performance
keywords: vbaac10.chm5187335
f1_keywords:
- vbaac10.chm5187335
ms.prod: access
ms.assetid: 122a8429-ad31-4e4b-2f68-d9d07c1deeeb
ms.date: 06/08/2017
---


# Improvements in Compilation Performance

Access includes improvements to module loading and compilation performance to make your code compile and run faster.


## Form and Report Modules Created on Demand

When you create a form or report in Access 2002 or later, the form or report doesn't automatically have an associated module. When you click  **Code** on the toolbar to view the form's or report's module, the module is created. You can also create the module from Visual Basic by referring to the form's **[Module](form-module-property-access.md)** property while the form or report is in Design view, or by setting the **[HasModule](form-hasmodule-property-access.md)** property to **True**.

The setting of the  **HasModule** property indicates whether a form or report currently has an associated module.

Since a form or report module isn't created until you need to add code to it, your project may have fewer modules to compile, resulting in improved compilation performance. Also, forms and reports without modules load more quickly than those with modules.


## Compiling on Demand

It's a good idea to explicitly compile the modules in your project by using the commands described above, but it's not necessary. Access compiles a module before running a procedure in it if the module hasn't already been compiled.

When a module is loaded for execution, Access checks to see whether the module has already been compiled. If not, Access compiles the module immediately prior to executing a procedure within it. The process of compiling slows down your code, so code in modules that have been saved in a compiled state will run faster.

Note that in Access 95, when you run a procedure in one module, all modules in the potential call tree are loaded, although by default they aren't compiled until a procedure within them is called. Because Access 97 (and later versions) load modules on a need-only basis, your code may run faster in many cases.

You can further enhance performance by grouping procedures in modules to reduce unnecessary compilations. Group procedures in modules with other procedures that they call, as opposed to grouping them in modules with unrelated procedures.


