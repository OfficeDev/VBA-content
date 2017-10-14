---
title: InvisibleApp.LanguageSettings Property (Visio)
keywords: vis_sdr.chm17560035
f1_keywords:
- vis_sdr.chm17560035
ms.prod: visio
api_name:
- Visio.InvisibleApp.LanguageSettings
ms.assetid: 0aff05cd-7655-0671-9c43-e45988c5a172
ms.date: 06/08/2017
---


# InvisibleApp.LanguageSettings Property (Visio)

Returns a reference to the Microsoft Office (MSO)  **LanguageSettings** interface. Read-only.


## Syntax

 _expression_ . **LanguageSettings**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Object


## Remarks

After you use the  **LanguageSettings** property to get a reference to the MSO **LanguageSettings** interface, you can use methods of that interface to get the locale identifier (LCID) for the language used when Office was installed, the user interface (UI) language, and the language for Help, as well as the current setting for the preferred language for editing in the UI.

However, you cannot use the  **LanguageSettings** interface to change language settings: you can change language settings only in the **Microsoft Office Language Settings 2007** dialog box. (Click **Start**, point to  **All Programs**, point to  **Microsoft Office**, point to  **Microsoft Office Tools**, and then click  **Microsoft Office 2007 Language Settings**. 


