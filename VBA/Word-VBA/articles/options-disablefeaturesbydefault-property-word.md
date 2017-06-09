---
title: Options.DisableFeaturesbyDefault Property (Word)
keywords: vbawd10.chm162988460
f1_keywords:
- vbawd10.chm162988460
ms.prod: word
api_name:
- Word.Options.DisableFeaturesbyDefault
ms.assetid: 58afcc8b-1d40-eebc-24ff-cb6bfdb5956d
ms.date: 06/08/2017
---


# Options.DisableFeaturesbyDefault Property (Word)

 **True** for Microsoft Word to disable in all documents all features introduced after the version of Word specified in the **[DisableFeaturesIntroducedAfterbyDefault](options-disablefeaturesintroducedafterbydefault-property-word.md)** . The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **DisableFeaturesbyDefault**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Remarks

The  **DisableFeaturesByDefault** property sets a global option for the application. If you want to disable features introduced after Word 97 for Windows for the document only, use the **[DisableFeatures](document-disablefeatures-property-word.md)** property.


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

