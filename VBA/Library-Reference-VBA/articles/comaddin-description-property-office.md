---
title: COMAddIn.Description Property (Office)
keywords: vbaof11.chm219001
f1_keywords:
- vbaof11.chm219001
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.COMAddIn.Description
ms.assetid: f194ae48-0762-732f-7c9a-f19a92e94d9b
---


# COMAddIn.Description Property (Office)

Gets or sets a descriptive  **String** value for the specified **COMAddin** object. Read/write.


## Syntax

 _expression_. **Description**

 _expression_ Required. A variable that represents a **[COMAddIn](comaddin-object-office.md)** object.


## Example

The following example displays the description text of the Microsoft Accessibility COM add-in for drawing.


```vb
MsgBox "The description of this " &; _ 
 "COMAddIn is """ &; Application.COMAddIns. _ 
 Item("msodraa9.ShapeSelect"). _ 
 Description &; """
```


## See also


#### Concepts


[COMAddIn Object](comaddin-object-office.md)

