---
title: Document.DisableFeaturesIntroducedAfter Property (Word)
keywords: vbawd10.chm158007639
f1_keywords:
- vbawd10.chm158007639
ms.prod: word
api_name:
- Word.Document.DisableFeaturesIntroducedAfter
ms.assetid: 5714062c-ffca-8feb-6b25-52f71568ae12
ms.date: 06/08/2017
---


# Document.DisableFeaturesIntroducedAfter Property (Word)

Disables all features introduced after a specified version of Microsoft Word in the document only. Read/write  **WdDisableFeaturesIntroducedAfter** .


## Syntax

 _expression_ . **DisableFeaturesIntroducedAfter**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **DisableFeatures** property must be set to **True** prior to setting the **DisableFeaturesIntroducedAfter** property. Otherwise, the setting will not take effect and will remain at its default setting of Word 97 for Windows.

The  **DisableFeaturesIntroducedAfter** property only affects the document for which the property is set. If you want to set a global option for the application to disable features for all documents, use the **DisableFeaturesIntroducedAfterByDefault** property.


## Example

This example disables all features added after Word for Windows 95, versions 7.0 and 7.0a, for the current document only. The global default setting remains unchanged.


```vb
Sub FeaturesDisable() 
 With ActiveDocument 
 
 'Checks whether features are disabled 
 If .DisableFeatures = True Then 
 
 'If they are, disables all features after Word for Windows 95 
 .DisableFeaturesIntroducedAfter = wd70 
 Else 
 
 'If not, turns on the disable features option and disables 
 'all features introduced after Word for Windows 95 
 .DisableFeatures = True 
 .DisableFeaturesIntroducedAfter = wd70 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

