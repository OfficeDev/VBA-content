---
title: Understanding the Lifetime of Variables
keywords: vbcn6.chm1076736
f1_keywords:
- vbcn6.chm1076736
ms.prod: office
ms.assetid: 018a61d5-4a0c-ac2e-6f06-50ba606034de
ms.date: 06/08/2017
---


# Understanding the Lifetime of Variables

The time during which a [variable](vbe-glossary.md) retains its value is known as its lifetime. The value of a variable may change over its lifetime, but it retains some value. When a variable loses [scope](vbe-glossary.md), it no longer has a value.

When a [procedure](vbe-glossary.md) begins running, all variables are initialized. A numeric variable is initialized to zero, a variable-length string is initialized to a zero-length string (""), and a fixed-length string is filled with the character represented by the ASCII character code 0, or **Chr(** 0 **)**. [Variant](vbe-glossary.md) variables are initialized to [Empty](vbe-glossary.md). Each element of a [user-defined type](vbe-glossary.md) variable is initialized as if it were a separate variable.

When you declare an [object variable](vbe-glossary.md), space is reserved in memory, but its value is set to  **Nothing** until you assign an object reference to it using the **Set** statement.

If the value of a variable isn't changed during the running of your code, it retains its initialized value until it loses scope.
A [procedure-level](vbe-glossary.md) variable declared with the **Dim** statement retains a value until the procedure is finished running. If the procedure calls other procedures, the variable retains its value while those procedures are running as well.
If a procedure-level variable is declared with the  **Static** keyword, the variable retains its value as long as code is running in any [module](vbe-glossary.md). When all code has finished running, the variable loses its scope and its value. Its lifetime is the same as a [module-level](vbe-glossary.md) variable.
A module-level variable differs from a static variable. In a [standard module](vbe-glossary.md) or a [class module](vbe-glossary.md), it retains its value until you stop running your code. In a class module, it retains its value as long as an instance of the class exists. Module-level variables consume memory resources until you reset their values, so use them only when necessary.
If you include the  **Static** keyword before a **Sub** or **Function** statement, the values of all the procedure-level variables in the procedure are preserved between calls.

