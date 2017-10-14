---
title: Ways to change the appearance of a control
keywords: fm20.chm5225237
f1_keywords:
- fm20.chm5225237
ms.prod: office
ms.assetid: b14bb419-dd2f-4f0b-9298-847082d93844
ms.date: 06/08/2017
---


# Ways to change the appearance of a control

Microsoft Forms includes several properties that let you define the appearance of controls in your application:



-  **ForeColor**
    
-  **BackColor**, **BackStyle**
    
-  **BorderColor**, **BorderStyle**
    
-  **SpecialEffect**
    

 **ForeColor** determines the[foreground color](glossary-vba.md). The foreground color applies to any text associated with the control, such as the caption or the control's contents.
 **BackColor** and **BackStyle** apply to the control's background. The background is the area within the control's boundaries, such as the area surrounding the text in a control, but not the control's border. **BackColor** determines the[background color](glossary-vba.md).  **BackStyle** determines whether the background is[transparent](glossary-vba.md). A transparent control background is useful if your application design includes a picture as the main background and you want to see that picture through the control.
 **BorderColor**, **BorderStyle**, and **SpecialEffect** apply to the control's border. You can use **BorderStyle** or **SpecialEffect** to choose a type of border. Only one of these two properties can be used at a time. When you assign a value to one of these properties, the system sets the other property to **None**. **SpecialEffect** lets you choose one of several border styles, but only lets you use[system colors](glossary-vba.md) for the border. **BorderStyle** supports only one border style, but lets you choose any color that is a valid setting for **BorderColor**. **BorderColor** specifies the color of the control's border, and is only valid when you use **BorderStyle** to create the border.

