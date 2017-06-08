---
title: Options.DisableFeaturesIntroducedAfterbyDefault Property (Word)
keywords: vbawd10.chm162988469
f1_keywords:
- vbawd10.chm162988469
ms.prod: word
api_name:
- Word.Options.DisableFeaturesIntroducedAfterbyDefault
ms.assetid: a7cf788b-f5c1-2d7e-b3de-1261b2a65c45
ms.date: 06/08/2017
---


# Options.DisableFeaturesIntroducedAfterbyDefault Property (Word)

Disables all features introduced after a the specified version for all documents. Read/write  **WdDisableFeaturesIntroducedAfter** .


## Syntax

 _expression_ . **DisableFeaturesIntroducedAfterbyDefault**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

The  **DisableFeaturesByDefault** property must be set to **True** prior to setting the **DisableFeaturesIntroducedAfterByDefault** property. Otherwise, the setting will not take effect and will remain at its default setting of Word 97 for Windows.

The  **DisableFeaturesIntroducedAfterByDefault** property sets a global option for the application and affects all documents. If you want to disable features introduced after a specified version for a document only, use the **DisableFeaturesIntroducedAfter** property.


## Example

This example disables all features introduced after Word for Windows 95, versions 7.0 and 7.0a, for all documents.


```vb
Sub FeaturesDisableByDefault() 
 With Application.Options 
 
 'Checks whether features are disabled 
 If .DisableFeaturesbyDefault = True Then 
 
 'If they are, disables all features after Word for Windows 95 
 .DisableFeaturesIntroducedAfterbyDefault = wd70 
 Else 
 
 'If not, turns on the disable features option and disables 
 'all features introduced after Word for Windows 95 
 .DisableFeaturesbyDefault = True 
 .DisableFeaturesIntroducedAfterbyDefault = wd70 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

