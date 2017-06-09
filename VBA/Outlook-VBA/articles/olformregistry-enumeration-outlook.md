---
title: OlFormRegistry Enumeration (Outlook)
keywords: vbaol11.chm3060
f1_keywords:
- vbaol11.chm3060
ms.prod: outlook
api_name:
- Outlook.OlFormRegistry
ms.assetid: 2d1076ae-0984-da03-a7ec-f083dc9d9e46
ms.date: 06/08/2017
---


# OlFormRegistry Enumeration (Outlook)

Indicates the form registry (library) where the  **Form** is stored.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olDefaultRegistry**|0|The Form is registered in the user's default form registry.|
| **olFolderRegistry**|3|The Form is registered in a form registry specific to a particular folder, and can only be accessed from that folder.|
| **olOrganizationRegistry**|4|The Form is registered in the organizational form registry. The form is available to all users.|
| **olPersonalRegistry**|2|The Form is registered in the user's personal registry and is only accessible to that user.|

## Remarks

Used as a parameter to the [FormDescription.PublishForm](formdescription-publishform-method-outlook.md) method to specify the form registry (library) in which to register the Form.


