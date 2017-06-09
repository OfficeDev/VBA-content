---
title: The code in this project must be updated for use on 64-bit systems
ms.prod: office
ms.assetid: 639e7a86-70c3-16dc-b73b-4f8ee82e816e
ms.date: 06/08/2017
---


# The code in this project must be updated for use on 64-bit systems

Complete error message:

The code in this project must be updated for use on 64-bit systems. Please review and update Declare statements and then mark them with the  **PtrSafe** attribute.

All  **[Declare Statements](declare-statement.md)** must now include the **[PtrSafe](ptrsafe-keyword.md)** keyword when running in 64-bit versions of Microsoft Office. The **PtrSafe** keyword indicates a **Declare** statement is safe to run in 64-bit versions of Microsoft Office.

Adding the  **PtrSafe** keyword to a **Declare** statement only signifies the **Declare** statement explicitly targets 64-bits, all data types within the statement that need to store 64-bits (including return values and parameters) must still be modified to hold 64-bit quantities using either[LongLong](longlong-data-type.md) for 64-bit integrals or[LongPtr](longptr-data-type.md) for pointers and handles.

