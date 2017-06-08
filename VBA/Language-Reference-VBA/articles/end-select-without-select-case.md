---
title: End Select without Select Case
keywords: vblr6.chm1011141
f1_keywords:
- vblr6.chm1011141
ms.prod: office
ms.assetid: 21fb3c2a-d273-1b2b-2ac2-e123fc3689ae
ms.date: 06/08/2017
---


# End Select without Select Case

 **End Select** must be matched with a preceding **Select Case**. This error has the following cause and solution:



- You used an  **End Select** statement without a corresponding **Select Case** statement. This is usually due to an extra **End Select** below a **Select Case** block, or leaving behind the **End Select** statement when copying a **Select Case** block from one[procedure](vbe-glossary.md) to another. Check each **End Select** statement to make sure it terminates a **Select Case** structure.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

