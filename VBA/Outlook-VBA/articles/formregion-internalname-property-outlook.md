---
title: FormRegion.InternalName Property (Outlook)
keywords: vbaol11.chm2400
f1_keywords:
- vbaol11.chm2400
ms.prod: outlook
api_name:
- Outlook.FormRegion.InternalName
ms.assetid: 2478d44e-887c-c245-6cfa-70a6a1e2c828
ms.date: 06/08/2017
---


# FormRegion.InternalName Property (Outlook)

Returns a  **String** that represents the internal programmatic name of the form region as defined in the manifest for the form region. Read-only.


## Syntax

 _expression_ . **InternalName**

 _expression_ A variable that represents a **FormRegion** object.


## Remarks

The internal name is required for a form region. The <name> tag in the corresponding form region manifest XML file maps to the value of the  **InternalName** property. For more information on the XML schema for form regions, see the Microsoft Outlook 2010 XML Schema Reference in the[MSDN Library](http://msdn.microsoft.com/library).

The value of the  **InternalName** property is used by the add-in or Microsoft Outlook to refer to the form region, for example, to determine which form region is being loaded or to load strings from the localized string resources. The **InternalName** property supports only ASCII characters. The string is case-insensitive, and its maximum length is 256 characters.


## See also


#### Concepts


[FormRegion Object](formregion-object-outlook.md)

