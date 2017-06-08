---
title: Transparency in Microsoft Forms
keywords: fm20.chm5225249
f1_keywords:
- fm20.chm5225249
ms.prod: office
ms.assetid: ce6991bf-f014-be02-971d-8c48e6cd52c1
ms.date: 06/08/2017
---


# Transparency in Microsoft Forms

Microsoft Forms supports transparency in two areas: the background of certain controls, and in bitmaps used on certain controls.

The  **BackStyle** property determines whether a control is[transparent](glossary-vba.md). A transparent control lets you see what is behind it on the form. This is useful if you have a decorative background on the form and you want to minimize the amount of that background that is hidden behind the controls. For more information on making a control transparent, see [Create a transparent control](create-a-transparent-control.md).

You can display a bitmap on many controls in Microsoft Forms. Certain controls support transparent bitmaps, that is, bitmaps in which one or more [background color](glossary-vba.md) is transparent. Bitmap transparency is not controlled by any control property; it is controlled by the color of the lower-left pixel in the image. Microsoft Forms does not provide a way to edit a bitmap and make it transparent; you must use a picture editor for this purpose.

In Microsoft Forms, bitmaps are always transparent on the following controls:


-  **CheckBox**
    
-  **CommandButton**
    
-  **Label**
    
-  **OptionButton**
    
-  **ToggleButton**
    

Transparent pictures sometimes have a hazy appearance. If you do not like this appearance, display the picture on a control that supports opaque images.
If you use a transparent bitmap on a control that does not support transparent bitmaps, the bitmap will display correctly but you won't be able to see what's behind it. In Microsoft Forms, the following controls do not support transparent bitmaps:


- The form window ( **UserForm** )
    
-  **Frame**
    
-  **Image**
    
-  **MultiPage**
    


