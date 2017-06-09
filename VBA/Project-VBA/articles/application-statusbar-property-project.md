---
title: Application.StatusBar Property (Project)
ms.prod: project-server
api_name:
- Project.Application.StatusBar
ms.assetid: c88965a0-302c-e0ce-ca5b-06fc2d21ff2d
ms.date: 06/08/2017
---


# Application.StatusBar Property (Project)

Gets or sets text in the status bar. Read/write  **Variant**.


## Syntax

 _expression_. **StatusBar**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **StatusBar** property returns **False** if the status bar is displaying the default text. Setting **StatusBar** to the Boolean value **False** restores the default text.


## Example

The following line of code sets custom text in the status bar.


```vb
Application.StatusBar = "This is custom text."
```


