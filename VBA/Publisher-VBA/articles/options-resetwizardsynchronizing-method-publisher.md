---
title: Options.ResetWizardSynchronizing Method (Publisher)
keywords: vbapb10.chm1048617
f1_keywords:
- vbapb10.chm1048617
ms.prod: publisher
api_name:
- Publisher.Options.ResetWizardSynchronizing
ms.assetid: 1027a113-45aa-b722-b625-a6bb7bbcc3e6
ms.date: 06/08/2017
---


# Options.ResetWizardSynchronizing Method (Publisher)

Resets the data that Microsoft Publisher uses to automatically change similar objects to have the same formatting or content.


## Syntax

 _expression_. **ResetWizardSynchronizing**

 _expression_A variable that represents an  **Options** object.


## Remarks

Unexpected formatting changes may be a result of Publisher's object synchronization. Resetting the synchronization data will stop these changes.


## Example

The following example resets the synchronization data that Publisher uses to give similar objects the same formatting.


```
Options.ResetWizardSynchronizing
```


