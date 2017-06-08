---
title: Page.Copy Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 6013fe1e-eb1c-dcca-b5eb-d99cc84f22fa
ms.date: 06/08/2017
---


# Page.Copy Method (Outlook Forms Script)

Copies the contents of an object to the Clipboard.


## Syntax

 _expression_. **Copy**

 _expression_A variable that represents a  **Page** object.


## Remarks

The original content remains on the object.

The actual content that is copied depends on the object. Using  **Copy** for a form, **[Frame](frame-object-outlook-forms-script.md)**, or  **[Page](page-object-outlook-forms-script.md)** copies the currently active control.


