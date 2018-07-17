---
title: TopLine Property (VBA Add-In Object Model)
keywords: vbob6.chm1071242
f1_keywords:
- vbob6.chm1071242
ms.prod: office
ms.assetid: 828ffefe-b76f-c58b-0558-c4e2b3f4c2e2
ms.date: 06/08/2017
---


# TopLine Property (VBA Add-In Object Model)



Returns a [Long](vbe-glossary.md) specifying the line number of the line at the top of the[code pane](vbe-glossary.md) or sets the line showing at the top of the code pane. Read/write.
 **Remarks**
Use the  **TopLine** property to return or set the line showing at the top of the code pane. For example, if you want line 25 to be the first line showing in a code pane, set the **TopLine** property to 25.
The  **TopLine** property setting must be a positive number. If the **TopLine** property setting is greater than the actual number of lines in the code pane, the setting will be the last line in the code pane.

