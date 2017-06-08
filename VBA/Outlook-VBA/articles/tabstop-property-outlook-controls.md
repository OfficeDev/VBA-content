---
title: TabStop Property (Outlook Controls)
keywords: olfm10.chm2002050
f1_keywords:
- olfm10.chm2002050
ms.prod: outlook
ms.assetid: a258b4c7-d388-9c92-c400-50bbdc023e9f
ms.date: 06/08/2017
---


# TabStop Property (Outlook Controls)

Returns or sets a  **Boolean** that indicates whether an object can receive focus when the user tabs to it. Read/write.


## Syntax

 _expression_. **TabStop**

 _expression_A variable that represents an Outlook control object.


## Remarks

 **True** to designate the object as a tab stop (default), **False** to bypass the object when the user is tabbing, although the object still holds its place in the actual tab order, as determined by the **[TabIndex](tabindex-property-outlook-controls.md)** property.

You can combine the settings of the  **Enabled** and the **TabStop** properties to prevent the user from selecting a command button with TAB, while still allowing the user to click the button. Setting **TabStop** to **False** means that the command button won't appear in the tab order.


