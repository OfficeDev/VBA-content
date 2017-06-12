---
title: Document.DisableFeatures Property (Word)
keywords: vbawd10.chm158007633
f1_keywords:
- vbawd10.chm158007633
ms.prod: word
api_name:
- Word.Document.DisableFeatures
ms.assetid: 40a62de3-f74e-d604-d3fc-dfb26abeb313
ms.date: 06/08/2017
---


# Document.DisableFeatures Property (Word)

 **True** disables all features introduced after the version specified in the **[DisableFeaturesIntroducedAfter](document-disablefeaturesintroducedafter-property-word.md)** property. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **DisableFeatures**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **DisableFeatures** property only affects the document for which you set the property. Use this property if you plan on sharing a document between users with earlier versions of Microsoft Word so that you don't end up introducing into a document features that are not available in their versions.


## Example

This example disables all features added after Word for Windows 95, versions 7.0 and 7.0a, for the current document. The global default setting remains unchanged.


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

