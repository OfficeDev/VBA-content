---
title: Form already displayed; can't show modally (Error 400)
keywords: vblr6.chm1117807
f1_keywords:
- vblr6.chm1117807
ms.prod: office
ms.assetid: 98f6191b-2756-4d5f-f9c3-47791b664cba
ms.date: 06/08/2017
---


# Form already displayed; can't show modally (Error 400)

You can't use the  **Show** method to display a visible form as modal. This error has the following cause and solution:



- You tried to use  **Show**, with the _style_[argument](vbe-glossary.md) set to 1 - **vbModal**, on an already visible form.
    
    Use either the  **Unload** statement or the **Hide** method on the form before trying to show it as a modal form.
    


