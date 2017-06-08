---
title: Application.Screen Property (Access)
keywords: vbaac10.chm12510
f1_keywords:
- vbaac10.chm12510
ms.prod: access
api_name:
- Access.Application.Screen
ms.assetid: d6faa33a-7701-d270-3bc7-04d53ac9303a
ms.date: 06/08/2017
---


# Application.Screen Property (Access)

You can use the  **Screen** property to return a reference the **[Screen](screen-object-access.md)** object and its related properties. Read-only.


## Syntax

 _expression_. **Screen**

 _expression_ A variable that represents an **Application** object.


## Remarks

 Use the **Screen** object to refer to a particular form, report, or control that has the focus.


## Example

The following example demonstrates how to change the cursor to an hourglass and back again to signify that some background activity is occurring.


```vb
Application.Screen.MousePointer = 11 ' Hourglass' Do some background activity.Application.Screen.MousePointer = 0 ' Back to normal
```


## See also


#### Concepts


[Application Object](application-object-access.md)

