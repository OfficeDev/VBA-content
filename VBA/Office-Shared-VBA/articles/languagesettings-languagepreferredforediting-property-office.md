---
title: LanguageSettings.LanguagePreferredForEditing Property (Office)
keywords: vbaof11.chm231002
f1_keywords:
- vbaof11.chm231002
ms.prod: office
api_name:
- Office.LanguageSettings.LanguagePreferredForEditing
ms.assetid: 345e29df-6cb7-13cc-a8ec-22196f38fc62
ms.date: 06/08/2017
---


# LanguageSettings.LanguagePreferredForEditing Property (Office)

Gets  **True** if the value for the **MsoLanguageID** constant has been identified in the Windows registry as a preferred language for editing. Read-only.


## Syntax

 _expression_. **LanguagePreferredForEditing**( **_lid_** )

 _expression_ A variable that represents a **LanguageSettings** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lid_|Required|**MsoLanguageID**|Returns one of the  **MsoLanguageID** enumerations.|

## Remarks

You must test all valid  **msoLanguageID** values to enumerate the set of preferred languages.


## Example

This example displays a message if U.S. English is a preferred editing language.


```
If Application.LanguageSettings. _ 
 LanguagePreferredForEditing(msoLanguageIDEnglishUS) Then 
 MsgBox "One of the preferred editing languages is US English." 
End If
```


## See also


#### Concepts


[LanguageSettings Object](languagesettings-object-office.md)
#### Other resources


[LanguageSettings Object Members](languagesettings-members-office.md)

