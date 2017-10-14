---
title: Change the Appearance of a Control
ms.prod: outlook
ms.assetid: 3e88980a-3b48-0ee9-06a7-f30aaf66f27b
ms.date: 06/08/2017
---


# Change the Appearance of a Control

Outlook includes several properties that you can use to define the appearance of controls in your form: 


-  **ForeColor** determines the foreground color. The foreground color applies to any text that is associated with the control, such as the caption or the control contents.
    
-  **BackColor** and **BackStyle** apply to the control background. The background is the area that is within the control boundaries, such as the area surrounding the text in a control, but not the control border. **BackColor** determines the background color. **BackStyle** determines whether the background is transparent. A transparent control background is useful if your form has a background picture. For **ForeColor** and **BackColor**, you can use the color scheme defined by your system, or you can use a custom color that you pick from the color palette. Using a system color, such as Menu Text, ensures that your form matches the colors and palette used by your applications. Custom colors do not always appear the same across systems and screen resolutions, but they do offer the widest choice of colors. 
    
-  **BorderColor**,  **BorderStyle**, and  **SpecialEffect** apply to the control border. You can use **BorderStyle** or **SpecialEffect** to choose a border type. Only one of these two properties can be used at a time. When you assign a value to one of these properties, the system sets the other property to **None**. With  **SpecialEffect**, you can choose one of several border styles, but you can only use system colors for the border.  **BorderStyle** supports only one border style, but you can choose any color that is a valid setting for **BorderColor**.  **BorderColor** specifies the color of the control border and is only valid when you use **BorderStyle** to create the border.
    
     **Note**  The  **BorderColor**,  **BorderStyle**, and  **SpecialEffect** properties can only be applied to the standard controls that are provided by default in the [Control Toolbox](control-toolbox-overview.md) and are not applicable to form regions.

Outlook supports transparency (that is, the display of whatever is behind an object instead of its background) in two areas: the background of certain controls and in bitmaps that are used on certain controls.

You can show a bitmap on many controls. Certain controls support transparent bitmaps — that is, bitmaps in which one or more background colors are transparent. Bitmap transparency is not controlled by any control property; it is controlled by the color of the lower-left pixel in the image. Outlook does not provide a way to edit a bitmap and make it transparent; instead, you must use a picture editor.
Bitmaps are always transparent on the following controls:  [CheckBox](checkbox-object-outlook-forms-script.md),  [CommandButton](commandbutton-object-outlook-forms-script.md),  [Label](label-object-outlook-forms-script.md),  [OptionButton](optionbutton-object-outlook-forms-script.md), and  [ToggleButton](togglebutton-object-outlook-forms-script.md). In Outlook, the following do not support transparent bitmaps: the form,  [Frame](frame-object-outlook-forms-script.md) control, [Image](image-object-outlook-forms-script.md) control, and [MultiPage](multipage-object-outlook-forms-script.md) control.
Transparent pictures sometimes have a hazy appearance. If you do not like that appearance, show the picture on a control that supports opaque images. If you use a transparent bitmap on a control that does not support transparent bitmaps, the bitmap appears correctly, but you cannot see what is behind the bitmap.
For more information, see the following topics:

-  [How to: Use a System Color for a Background or Foreground](use-a-system-color-for-a-background-or-foreground.md)
    
-  [How to: Use a Custom Color for the Background or Foreground of a Control](use-a-custom-color-for-the-background-or-foreground-of-a-control.md)
    
-  [ How to: Set the Background Color of a Form](set-the-background-color-of-a-form.md)
    
-  [How to: Make a Control Transparent](make-a-control-transparent.md)
    

