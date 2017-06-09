---
title: ExchangeUser.Details Method (Outlook)
keywords: vbaol11.chm2074
f1_keywords:
- vbaol11.chm2074
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Details
ms.assetid: 6c93a583-cc61-e527-7832-88dba525854a
ms.date: 06/08/2017
---


# ExchangeUser.Details Method (Outlook)

Displays a modal dialog box that provides detailed information about an  **[ExchangeUser](exchangeuser-object-outlook.md)** object.


## Syntax

 _expression_ . **Details**( **_HWnd_** )

 _expression_ A variable that represents an **ExchangeUser** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HWnd_|Optional| **Variant**| The parent window handle for the Details dialog box. A zero value (the default) specifies a modal dialog box.|

## Remarks

The  **Details** method fails if the **[ExchangeUser.Name](exchangeuser-name-property-outlook.md)** property is empty. You must use error handling to handle run-time errors, and when the user clicks **Cancel** in the dialog box.

The  **Details** method actually stops the code from running while the dialog box is displayed.


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

