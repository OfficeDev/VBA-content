---
title: Close Method (VBA Add-In Object Model)
keywords: vbob6.chm100111
f1_keywords:
- vbob6.chm100111
ms.prod: MULTIPLEPRODUCTS
ms.assetid: e3c951ed-032b-9e4b-ba1b-a802f42d3544
---


# Close Method (VBA Add-In Object Model)



Closes and destroys a window.
 **Syntax**
 _object_**.Close**
The  _object_ placeholder is an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.
 **Remarks**
The following types of windows respond to the  **Close** method in different ways:


- For a window that is a [code pane](vbe-glossary.md),  **Close** destroys the code pane.
    
- For a window that is a [designer](vbe-glossary.md),  **Close** destroys the contained designer.
    
- For windows that are always available on the  **View** menu, **Close** hides the window.
    


