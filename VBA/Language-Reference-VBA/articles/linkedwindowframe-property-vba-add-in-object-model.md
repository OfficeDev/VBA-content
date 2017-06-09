---
title: LinkedWindowFrame Property (VBA Add-In Object Model)
keywords: vbob6.chm1071224
f1_keywords:
- vbob6.chm1071224
ms.prod: office
ms.assetid: d97711c2-50e5-583f-70a1-ec25b0e1999f
ms.date: 06/08/2017
---


# LinkedWindowFrame Property (VBA Add-In Object Model)



Returns the  **Window** object representing the frame that contains the window. Read-only.
 **Remarks**
The  **LinkedWindowFrame** property enables you to access the object representing the[linked window frame](vbe-glossary.md), which has properties distinct from the window or windows it contains. If the window isn't linked, the  **LinkedWindowFrame** property returns **Nothing**.


 **Important**  Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.



