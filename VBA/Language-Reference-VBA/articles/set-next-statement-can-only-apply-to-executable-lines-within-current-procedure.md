---
title: Set Next Statement can only apply to executable lines within current procedure.
keywords: vblr6.chm1040350
f1_keywords:
- vblr6.chm1040350
ms.prod: office
ms.assetid: 4e5c0a9d-95ec-ee89-499f-42af2b9d44ec
ms.date: 06/08/2017
---


# Set Next Statement can only apply to executable lines within current procedure.

You can choose the  **Set Next Statement** command to indicate where a suspended program should begin when execution is continued. This error has the following causes and solutions:



- When you chose the command, the cursor was on a line that didn't contain an executable statement. Place the cursor on a line with an executable statement and try again. [Declarations](vbe-glossary.md), [line labels](vbe-glossary.md), and [comments](vbe-glossary.md) aren't executable, so lines with only declarations, labels, and comments can't be the targets of the **Set Next Statement** command.
    
- When you chose the command, the cursor was on a line outside the currently executing [procedure](vbe-glossary.md).
    
    Place the cursor on a line within the currently executing procedure.
    


