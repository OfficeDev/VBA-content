---
title: Create a Set of OptionButton Controls
ms.prod: outlook
ms.assetid: 6aee3c64-df73-df1a-0db8-2740f8dec0d9
ms.date: 06/08/2017
---


# Create a Set of OptionButton Controls

By default, all  [OptionButton](optionbutton-object-outlook-forms-script.md)controls in a container are part of a single option group. This means that selecting one of the buttons automatically sets all other option buttons on the form to  **False**.

If you want more than one option group on the form, there are two ways to create additional groups:

- Use the  [GroupName](optionbutton-groupname-property-outlook-forms-script.md)property to identify related buttons. This method reduces the number of controls required on the form, which can reduce the hard disk space required and improve the performance of the form. If you want to create an option group in a  [TabStrip](tabstrip-object-outlook-forms-script.md)(which is not a container), you must use the  **GroupName** property. For more information, see [How to: Create a Set of OptionButtons by Using the GroupName Property](create-a-set-of-optionbuttons-by-using-the-groupname-property.md).
    
- Put related buttons in a  **[Page](page-object-outlook-forms-script.md)**,  **[MultiPage](multipage-object-outlook-forms-script.md)**, or  **[Frame](frame-object-outlook-forms-script.md)** on the form. For more information, see  [How to: Add a Control to a Form](add-a-control-to-a-form.md).
    

