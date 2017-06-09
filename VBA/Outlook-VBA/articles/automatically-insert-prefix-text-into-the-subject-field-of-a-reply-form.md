---
title: Automatically Insert Prefix Text into the Subject Field of a Reply Form
ms.prod: outlook
ms.assetid: 8e35cbd6-1ce2-7a73-4365-9082b1c745e1
ms.date: 06/08/2017
---


# Automatically Insert Prefix Text into the Subject Field of a Reply Form

## Customizing with form regions

In a form region, you can set the subject prefix in the XML manifest file of the form region. Alternatively, you can write code in the  [BeforeFormRegionShow](formregionstartup-beforeformregionshow-method-outlook.md) method of the [FormRegionStartup](formregionstartup-object-outlook.md) interface.

For more information, see  [How to: Specify a Subject Prefix of an Item Resulting from an Action](specify-a-subject-prefix-of-an-item-resulting-from-an-action.md).


## Customizing with form pages


1. In the Forms Designer, click the  **(Actions)** page of your form.
    
2. Click the action to which you want to add the prefix text, and then click  **Properties**. 
    
3. In the  **Subject prefix** box, type the text as you want it to appear (Outlook automatically adds a colon after the text).
    
    For example, you could add the prefix "Re:" to a reply form.
    

