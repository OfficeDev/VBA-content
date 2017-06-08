---
title: Can't load or unload this object (Error 361)
keywords: vblr6.chm1117809
f1_keywords:
- vblr6.chm1117809
ms.prod: office
ms.assetid: 78438f88-b013-a3d1-9a57-f3a1781691f5
ms.date: 06/08/2017
---


# Can't load or unload this object (Error 361)

A  **Load** statement or **Unload** statement has referenced an invalid object or control. This error has the following causes and solutions:



- You tried to load or unload an object that isn't a loadable object, such as Screen, Printer, or Clipboard. Delete the erroneous statement from your code.
    
- You tried to load or unload an existing control that isn't part of a [control array](vbe-glossary.md). For example, assuming that a  **TextBox** with the **Name** property Text1 exists, `Load Text1`will cause this error.
    
    Delete the erroneous statement from your code or change the reference to a control in a control array.
    
- You tried to unload a  **Menu** control in the Click event of its parent menu. Unload the **Menu** control with some other[procedure](vbe-glossary.md).
    
- You tried to unload the last visible menu item of a  **Menu** control. Check the setting of the **Visible** property for the other menu items in the control array before trying to unload a menu item, or delete the erroneous statement from your code.
    


