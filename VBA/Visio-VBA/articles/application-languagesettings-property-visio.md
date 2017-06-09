---
title: Application.LanguageSettings Property (Visio)
keywords: vis_sdr.chm10060035
f1_keywords:
- vis_sdr.chm10060035
ms.prod: visio
api_name:
- Visio.Application.LanguageSettings
ms.assetid: 3fa0c4a4-3a1c-b035-9f9d-e4358917ebee
ms.date: 06/08/2017
---


# Application.LanguageSettings Property (Visio)

Returns a reference to the Microsoft Office (MSO)  **LanguageSettings** interface. Read-only.


## Syntax

 _expression_ . **LanguageSettings**

 _expression_ A variable that represents an **Application** object.


### Return Value

Object


## Remarks

After you use the  **LanguageSettings** property to get a reference to the MSO **LanguageSettings** interface, you can use methods of that interface to get the locale identifier (LCID) for the language used when Office was installed, the user interface (UI) language, and the language for Help, as well as the current setting for the preferred language for editing in the UI, as shown in the following example.

However, you cannot use the  **LanguageSettings** interface to change language settings: you can change language settings only in the **Microsoft Office Language Settings 2007** dialog box. (Click **Start**, point to  **All Programs**, point to  **Microsoft Office**, point to  **Microsoft Office Tools**, and then click  **Microsoft Office 2007 Language Settings**. 


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **LanguageSettings** property to get an MSO **LanguageSettings** interface, and then to use two of its methods to get the ID of the language set for the UI, and to test whether US English is set as the preferred language for editing.


```vb
Public Sub LanguageSettings_Example() 
 
    Dim msoLanguageSettings As LanguageSettings 
 
    Set msoLanguageSettings = Application.LanguageSettings 
    Debug.Print msoLanguageSettings.LanguageID(msoLanguageIDUI) 
    Debug.Print msoLanguageSettings.LanguagePreferredForEditing(msoLanguageIDEnglishUS) 
     
End Sub
```


