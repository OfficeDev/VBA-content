---
title: Inspectors.NewInspector Event (Outlook)
keywords: vbaol11.chm312
f1_keywords:
- vbaol11.chm312
ms.prod: outlook
api_name:
- Outlook.Inspectors.NewInspector
ms.assetid: 945fb1a6-262f-da0d-16c6-bc27193505ac
ms.date: 06/08/2017
---


# Inspectors.NewInspector Event (Outlook)

Occurs whenever a new inspector window is opened, either as a result of user action or through program code. 


## Syntax

 _expression_ . **NewInspector**( **_Inspector_** )

 _expression_ A variable that represents an **Inspectors** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Inspector_|Required| **[Inspector](inspector-object-outlook.md)**|The inspector that was opened.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

The event occurs after the new  **Inspector** object is created but before the inspector window appears.


## See also


#### Concepts


[Inspectors Object](inspectors-object-outlook.md)

