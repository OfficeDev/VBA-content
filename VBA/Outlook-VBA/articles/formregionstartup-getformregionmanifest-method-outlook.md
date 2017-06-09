---
title: FormRegionStartup.GetFormRegionManifest Method (Outlook)
keywords: vbaol11.chm3305
f1_keywords:
- vbaol11.chm3305
ms.prod: outlook
api_name:
- Outlook.FormRegionStartup.GetFormRegionManifest
ms.assetid: de752c6f-423a-ee2f-aa7e-d1107cf406a2
ms.date: 06/08/2017
---


# FormRegionStartup.GetFormRegionManifest Method (Outlook)

Obtains the XML manifest for a form region.


## Syntax

 _expression_ . **GetFormRegionManifest**( **_FormRegionName_** , **_LCID_** )

 _expression_ A variable that represents a **FormRegionStartup** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FormRegionName_|Required| **String**|The name of the form region which is the name used when registering the form region in the Windows registry.|
| _LCID_|Required| **Long**|The locale ID that identifies the language that Outlook is currently using. This value is used to obtain the localization strings corresponding to this language for the form region.|

### Return Value

A  **Variant** that represents the XML manifest for a form region. This XML string includes characteristics of the form region such as the display name (as specified by the title element), any associated layout file or add-in, any supported user actions, and any localization strings. The XML must follow the form region XML schema. For more information on the form region XML schema, see the Microsoft Outlook 2010 XML Schema Reference in the[MSDN Library](http://msdn.microsoft.com/library).


## Remarks

This method is intended to be implemented by an add-in and called by Outlook. As part of the  **[FormRegionStartup](formregionstartup-object-outlook.md)** interface, this method and the **[GetFormRegionIcon](formregionstartup-getformregionicon-method-outlook.md)** method provide a mechanism through which an add-in can register a form region and provide Outlook the XML manifest and the icons for the form region.

If you would like an add-in to provide the XML manifest for a form region, specify the  **ProgID** of the add-in when you register the form region in the Windows registry. For more information on registering a form region, see[Specifying Form Regions in the Windows Registry](http://msdn.microsoft.com/library/0de3fcb1-b357-8300-c943-9a5a788d4976%28Office.15%29.aspx). The add-in must implement the  **GetFormRegionManifest** method of the **FormRegionStartup** interface. Note that if you do not specify any **ProgID** in the Windows registry, Outlook will not call this method.

Relying on an add-in to provide the XML manifest for a form region also means you are allowing the add-in to provide any icons for the form region. The add-in must also implement the  **GetFormRegionIcon** of the **FormRegionStartup** interface. Outlook will call **GetFormRegionIcon** to obtain any add-in specified icons for the form region. For more information on using an add-in to specify icons, see[How to: Use an Add-in to Specify Icons for a Form Region](http://msdn.microsoft.com/library/7d542c9b-1881-780a-b58d-e34639399b60%28Office.15%29.aspx).

 When Outlook starts, it reads the list of form regions from the Windows registry and caches the data. Based on this data, if Outlook notices that an add-in has been specified to provide the XML manifest for a form region, Outlook will use the **ProgID** provided in the cached data and call the **GetFormRegionManifest** method implemented by this add-in to obtain the XML it needs to display the form region. If the XML manifest is not valid and does not conform to the form region XML schema, Outlook will not be able to load the form region. Also, if you do not specify any **ProgID** in the Windows registry, Outlook will not call the **GetFormRegionManifest** and **GetFormRegionIcon** methods.

Outlook ignores the following elements when the add-in provides the XML manifest: 


-  **name** : Outlook ignores the value specified for this element and will use the name specified for the form region in the registry.
    
-  **layoutFile** : Outlook ignores this element because an add-in is extending this form region.
    
-  **addin** : Outlook uses the value that is preceded by an equal sign ( **=**) in the registry as the  **ProgID** of the add-in.
    
-  **file** attribute of **stringOverride** : Outlook ignores any secondary localization file specified by the **stringOverride** element. The add-in can implement **GetFormRegionManifest** to return inline the XML manifest for string localization for the specified _LCID_ , or manage string localization in another way, for example, using .NET Framework localization, and then return the appropriate XML manifest for the specified _LCID_ .
    



## See also


#### Concepts


[FormRegionStartup Interface](formregionstartup-object-outlook.md)

