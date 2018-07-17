---
title: Application.WebHideToolbars Method (Project)
keywords: vbapj.chm1306
f1_keywords:
- vbapj.chm1306
ms.prod: project-server
api_name:
- Project.Application.WebHideToolbars
ms.assetid: c6e323c9-b1a4-79bb-d714-b7ddaebbf619
ms.date: 06/08/2017
---


# Application.WebHideToolbars Method (Project)

Shows or hides all toolbars except the  **Menu** and **Web** toolbars. Obsolete in Project.


## Syntax

 _expression_. **WebHideToolbars**( ** _Hide_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Hide_|Optional|**Boolean**|**True** if all toolbars except the **Menu** and **Web** toolbars are hidden. The default value is **True** if toolbars other than **Menu** and **Web** are displayed, and **False** if they are not.|

### Return Value

 **Boolean**


### Remarks

Project does not use toolbars.


