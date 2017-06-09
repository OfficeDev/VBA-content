---
title: Must close or hide topmost modal form first (Error 402)
keywords: vblr6.chm402
f1_keywords:
- vblr6.chm402
ms.prod: office
ms.assetid: cba9d4b4-f8d9-0ba5-340f-38bd16cc59d7
ms.date: 06/08/2017
---


# Must close or hide topmost modal form first (Error 402)

The modal form you are trying to close or hide isn't on top of the [z-order](vbe-glossary.md). This error has the following cause and solution:



- Another modal form is higher in the z-order than the modal form you tried to close or hide. First use either the  **Unload** statement or the **Hide** method on any modal form higher in the z-order. A modal form is a form displayed by the **Show** method, with the _style_ argument set to 1 - **vbModal**.
    


