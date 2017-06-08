---
title: AddressEntry.Details Method (Outlook)
keywords: vbaol11.chm2051
f1_keywords:
- vbaol11.chm2051
ms.prod: outlook
api_name:
- Outlook.AddressEntry.Details
ms.assetid: 85457da6-c97a-387d-6c7e-40eb005b25aa
ms.date: 06/08/2017
---


# AddressEntry.Details Method (Outlook)

Displays a modeless dialog box that provides detailed information about an  **[AddressEntry](addressentry-object-outlook.md)** object.


## Syntax

 _expression_ . **Details**( **_HWnd_** )

 _expression_ An expression that returns a **AddressEntry** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HWnd_|Optional| **Variant**|The parent window handle for the  **Details** dialog box. A zero value (the default) specifies that the dialog is parented to Outlook.|

## Remarks


 **Note**  The  **Details** method fails if the **[Name](addressentry-name-property-outlook.md)** property is empty.

You must use error handling to handle run-time errors when the user clicks  **Cancel** in the dialog box. The **Details** method actually stops the code from running while the dialog box is displayed.


## See also


#### Concepts


[AddressEntry Object](addressentry-object-outlook.md)

