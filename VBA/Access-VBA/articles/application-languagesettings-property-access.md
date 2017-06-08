---
title: Application.LanguageSettings Property (Access)
keywords: vbaac10.chm12588
f1_keywords:
- vbaac10.chm12588
ms.prod: access
api_name:
- Access.Application.LanguageSettings
ms.assetid: f2b039bf-95a8-7820-355e-67fa5e47aaf6
ms.date: 06/08/2017
---


# Application.LanguageSettings Property (Access)

You can use the  **LanguageSettings** property to return a read-only reference to the current **LanguageSettings** object and its related properties.


## Syntax

 _expression_. **LanguageSettings**

 _expression_ A variable that represents an **Application** object.


## Remarks

Once you establish a reference to the  **LanguageSettings** object, you can access all the properties and methods of the object. You can set a reference to the **LanguageSettings** object by clicking **References** on the **Tools** menu while in module Design view. Then set a reference to the Microsoft Office Object Library in the **References** dialog box by selecting the appropriate check box. Microsoft Access can set this reference for you if you use a Microsoft Office Object Library constant to set a **LanguageSettings** object's property or as an argument to a **LanguageSettings** object's method.


## Example

The following example displays a message indicating the language Access uses for Help on the user's machine. A listing of all the available languages and their identification numbers is available in the Visual Basic Editor by selecting  **Object Browser** from the **View** menu, typing the word **MsoLanguageID** in the **Search Text box**, and clicking the **Search** button.


```vb
Dim mli As MsoLanguageID 
mli = Application.LanguageSettings.LanguageID(msoLanguageIDHelp) 
MsgBox "The language ID used for Access Help is " &; mli
```


## See also


#### Concepts


[Application Object](application-object-access.md)

