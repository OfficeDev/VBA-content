---
title: Resume without error (Error 20)
keywords: vblr6.chm1011265
f1_keywords:
- vblr6.chm1011265
ms.prod: office
ms.assetid: 02b7eb1c-a637-810d-78fd-1945a5784a54
ms.date: 06/08/2017
---


# Resume without error (Error 20)

A  **Resume** statement can only appear within an error handler and can only be executed in an active error handler. This error has the following causes and solutions:



- You placed a  **Resume** statement outside error-handling code. Move the statement into an error handler, or delete it.
    
- Your code jumped into an error handler even though there was no error. Perhaps you misspelled a [line label](vbe-glossary.md). Jumps to labels can't occur across [procedures](vbe-glossary.md), so search the procedure for the label that identifies the error handler. If you find a duplicate label specified as the target of a  **GoTo** statement that isn't an **On Error GoTo** statement, change the line label to agree with its intended target.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

