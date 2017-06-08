---
title: Form.LayoutForPrint Property (Access)
keywords: vbaac10.chm13391
f1_keywords:
- vbaac10.chm13391
ms.prod: access
api_name:
- Access.Form.LayoutForPrint
ms.assetid: fd8c8112-186a-3f77-06ef-783bf48a7052
ms.date: 06/08/2017
---


# Form.LayoutForPrint Property (Access)

You can use the  **LayoutForPrint** property to specify whether the form uses printer or screen fonts. Read/write **Boolean**.


## Syntax

 _expression_. **LayoutForPrint**

 _expression_ A variable that represents a **Form** object.


## Remarks

When you choose a font in Microsoft Access, you are choosing either a screen font or a printer font, depending on the setting of the  **LayoutForPrint** property. Remember that printer fonts and screen fonts can differ, and characters on screen may not look exactly like those displayed on the printed page.

Screen fonts are the images of letters, numbers, and symbols that are installed on your system to be displayed on the screen. If you installed a printer, additional screen fonts may have been installed automatically.

Printer fonts are the letters, numbers, and symbols that are produced when you print a form. The available fonts are those fonts that were installed as part of your printer's setup, and depend on your printer.

If you design a form on a system with a different printer than the one you will use to print, Microsoft Access displays a message when you print the form to let you know that it was designed for another kind of printer. If you print the form anyway, your printer may substitute different fonts. Similarly, Microsoft Access may substitute fonts if you change the  **LayoutForPrint** property setting. For example, you might design a form with **LayoutForPrint** set to No, then change the setting to Yes. You can reselect the font for each control to specify the appearance of the form.


## See also


#### Concepts


[Form Object](form-object-access.md)

