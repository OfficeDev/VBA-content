---
title: OlkLabel.UseHeaderColor Property (Outlook)
keywords: vbaol11.chm1000497
f1_keywords:
- vbaol11.chm1000497
ms.prod: outlook
api_name:
- Outlook.OlkLabel.UseHeaderColor
ms.assetid: 9b205ce8-0875-06da-6746-641ae889d4df
ms.date: 06/08/2017
---


# OlkLabel.UseHeaderColor Property (Outlook)

Returns or sets a  **Boolean** that indicates whether the label control should use the foreground color to match the current Windows XP or Windows Vista theme. Read/write.


## Syntax

 _expression_ . **UseHeaderColor**

 _expression_ A variable that represents an **olkLabel** object.


## Remarks

This property is intended for label controls in a message form displayed in the Reading Pane and in the Inspector. If the property is  **True** , then the label should use the foreground color that matches the current Windows theme. If the property is **False** , then the label should use the foreground color indicated by the **[ForeColor](olklabel-forecolor-property-outlook.md)** property.


## See also


#### Concepts


[OlkLabel Object](olklabel-object-outlook.md)

