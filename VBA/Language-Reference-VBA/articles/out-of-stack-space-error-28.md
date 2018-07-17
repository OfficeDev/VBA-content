---
title: Out of stack space (Error 28)
keywords: vblr6.chm1000028
f1_keywords:
- vblr6.chm1000028
ms.prod: office
ms.assetid: ce345551-ad57-1120-546a-239d144c330a
ms.date: 06/08/2017
---


# Out of stack space (Error 28)

The stack is a working area of memory that grows and shrinks dynamically with the demands of your executing program. This error has the following causes and solutions:



- You have too many active  **Function**, **Sub**, or **Property** procedure calls. Check that[procedures](vbe-glossary.md) aren't nested too deeply. This is especially true with recursive procedures, that is, procedures that call themselves. Make sure recursive procedures terminate properly. Use the **Calls** dialog box to view which procedures are active (on the stack).
    
- Your local [variables](vbe-glossary.md) require more local variable space than is available.
    
    Try declaring some variables at the [module level](vbe-glossary.md) instead. You can also declare all variables in the procedure static by preceding the **Property**, **Sub**, or **Function** keyword with **Static**. Or you can use the **Static** statement to declare individual **Static** variables within procedures.
    
- You have too many fixed-length strings. Fixed-length strings in a procedure are more quickly accessed, but use more stack space than variable-length strings, because the string data itself is placed on the stack. Try redefining some of your fixed-length strings as variable-length strings. When you declare variable-length strings in a procedure, only the string descriptor (not the data itself) is placed on the stack. You can also define the string at module level where it requires no stack space. Variables declared at module level are  **Public** by default, so the string is visible to all procedures in the module.
    
- You have too many nested  **DoEvents** function calls. Use the **Calls** dialog box to view which procedures are active on the stack.
    
- Your code triggered an event cascade. An event cascade is caused by triggering an event that calls an event procedure that's already on the stack. An event cascade is similar to an unterminated recursive procedure call, but it's less obvious, since the call is made by Visual Basic rather than by an explicit call in your code. Use the  **Calls** dialog box to view which procedures are active (on the stack).
    

To display the  **Calls** dialog box, select the **Calls** button to the right of the[Procedure box](vbe-glossary.md) in the **Debug** window or choose the **Calls** command. For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

