---
title: Invalid procedure call or argument (Error 5)
keywords: vblr6.chm1000005
f1_keywords:
- vblr6.chm1000005
ms.prod: office
ms.assetid: 481b8431-b4ba-b368-2c5e-ade85b99348d
ms.date: 06/08/2017
---


# Invalid procedure call or argument (Error 5)

Some part of the call can't be completed. This error has the following causes and solutions:



- An [argument](vbe-glossary.md) probably exceeds the range of permitted values. For example, the **Sin** function can only accept values within a certain range. Positive arguments less than 2,147,483,648 are accepted, while 2,147,483,648 generates this error.
    
    Check the ranges permitted for arguments.
    
- This error can also occur if an attempt is made to call a [procedure](vbe-glossary.md) that isn't valid on the current platform. For example, some procedures may only be valid for Microsoft Windows, or for the Macintosh, and so on.
    
    Check platform-specific information about the procedure.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

