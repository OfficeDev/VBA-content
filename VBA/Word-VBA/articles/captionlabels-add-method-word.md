---
title: CaptionLabels.Add Method (Word)
keywords: vbawd10.chm158859364
f1_keywords:
- vbawd10.chm158859364
ms.prod: word
api_name:
- Word.CaptionLabels.Add
ms.assetid: f74af8c0-fa16-8ea2-3012-ac207d187502
ms.date: 06/08/2017
---


# CaptionLabels.Add Method (Word)

Returns a  **CaptionLabel** object that represents a custom caption label.


## Syntax

 _expression_ . **Add**( **_Name_** )

 _expression_ Required. A variable that represents a **[CaptionLabels](captionlabels-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the custom caption label.|

### Return Value

CaptionLabel


## Example

This example adds a custom caption label named Demo Slide. To verify that the custom label is added, view the  **Label** combo box in the **Caption** dialog box, accessed from the **Reference** command on the **Insert** menu.


```vb
Sub CapLbl() 
 CaptionLabels.Add Name:="Demo Slide" 
End Sub
```


## See also


#### Concepts


[CaptionLabels Collection Object](captionlabels-object-word.md)

