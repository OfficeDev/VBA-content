---
title: NavigationGroup.Name Property (Outlook)
keywords: vbaol11.chm2888
f1_keywords:
- vbaol11.chm2888
ms.prod: outlook
api_name:
- Outlook.NavigationGroup.Name
ms.assetid: ad66ef0a-1348-372a-f98a-d43171856b35
ms.date: 06/08/2017
---


# NavigationGroup.Name Property (Outlook)

Returns or sets a  **String** value that represents the display name for the **[NavigationGroup](navigationgroup-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **NavigationGroup** object.


## Remarks

This property is limited to 127 characters.  **String** values longer than 127 characters are truncated.

An error occurs if you attempt to set the value of this property for any  **NavigationGroup** object associated with a **[MailModule](mailmodule-object-outlook.md)** object.


## See also


#### Concepts


[NavigationGroup Object](navigationgroup-object-outlook.md)

