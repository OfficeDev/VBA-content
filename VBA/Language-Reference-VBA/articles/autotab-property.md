---
title: AutoTab Property
keywords: fm20.chm2000750
f1_keywords:
- fm20.chm2000750
ms.prod: office
api_name:
- Office.AutoTab
ms.assetid: 36af6755-72d8-439a-2999-fc4224760529
ms.date: 06/08/2017
---


# AutoTab Property



Specifies whether an automatic tab occurs when a user enters the maximum allowable number of characters into a  **TextBox** or the text box portion of a **ComboBox**.
 **Syntax**
 _object_. **AutoTab** [= _Boolean_ ]
The  **AutoTab** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Specifies whether an automatic tab occurs.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|Tab occurs.|
|**False**|Tab does not occur (default).|
 **Remarks**
The  **MaxLength** property specifies the maximum number of characters allowed in a **TextBox** or the text box portion of a **ComboBox**.
You can specify the  **AutoTab** property for a **TextBox** or **ComboBox** on a form for which you usually enter a set number of characters. Once a user enters the maximum number of characters, the[focus](vbe-glossary.md) automatically moves to the next control in the[tab order](vbe-glossary.md). For example, if a  **TextBox** displays inventory stock numbers that are always five characters long, you can use **MaxLength** to specify the maximum number of characters to enter into the **TextBox** and **AutoTab** to automatically tab to the next control after the user enters five characters.
Support for  **AutoTab** varies from one application to another. Not all[containers](vbe-glossary.md) support this property.

