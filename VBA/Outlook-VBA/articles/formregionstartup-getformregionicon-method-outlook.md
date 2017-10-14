---
title: FormRegionStartup.GetFormRegionIcon Method (Outlook)
keywords: vbaol11.chm3307
f1_keywords:
- vbaol11.chm3307
ms.prod: outlook
api_name:
- Outlook.FormRegionStartup.GetFormRegionIcon
ms.assetid: c1c0bd3f-3fae-8e9b-d579-58d609bbaa4e
ms.date: 06/08/2017
---


# FormRegionStartup.GetFormRegionIcon Method (Outlook)

Obtains an icon image that will be displayed for a particular type of icon for the form region.


## Syntax

 _expression_ . **GetFormRegionIcon**( **_FormRegionName_** , **_LCID_** , **_Icon_** )

 _expression_ A variable that represents a **FormRegionStartup** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FormRegionName_|Required| **String**|The name of the form region which is the name used when registering the form region in the Windows registry.|
| _LCID_|Required| **Long**|The locale ID that identifies the language that Outlook is currently using. This value is used to obtain the localization strings corresponding to this language for the form region.|
| _Icon_|Required| **[OlFormRegionIcon](olformregionicon-enumeration-outlook.md)**|A constant that identifies the type of icon.|

### Return Value

A Variant that is either a byte-array that represents the original bytes of the image file or an  **IPictureDisp** object.


## Remarks

This method is intended to be implemented by an add-in and called by Outlook. As part of the  **[FormRegionStartup](formregionstartup-object-outlook.md)** interface, this method and the **[GetFormRegionManifest](formregionstartup-getformregionmanifest-method-outlook.md)** method provide a mechanism through which an add-in can register a form region and provide Outlook with the XML manifest and the icons for the form region.

If you would like an add-in to provide icons for a form region, specify the ProgID of the add-in when you register the form region in the Windows registry. For more information on registering a form region, see [Specifying Form Regions in the Windows Registry](http://msdn.microsoft.com/library/0de3fcb1-b357-8300-c943-9a5a788d4976%28Office.15%29.aspx). The add-in must implement the  **GetFormRegionManifest** and the **GetFormRegionIcon** methods of the **FormRegionStartup** interface.

In the XML manifest for the form region, under the  **icons** element, specify the value `addin` for each of the child elements where you would like to use a custom icon. Implement **GetFormRegionIcon** such that when Outlook passes that type of icon as an argument for _Icon_ , **GetFormRegionIcon** returns the image of the custom icon. If you want Outlook to display the default icon, implement **GetFormRegionIcon** such that it returns **null** ( **Nothing** in Visual Basic) for that type of icon. **GetFormRegionIcon** should also return **null** ( **Nothing** in Visual Basic) when _Icon_ is **olFormRegionIconDefault** .

 When Outlook starts, it reads the list of form regions from the Windows registry and caches the data associated with the form regions. If a form region has been registered with a ProgID, Outlook will resort to the corresponding add-in by calling its implementation of **GetFormRegionIcon** for any icon in the XML manifest that has `addin` as the value of a child element of the **icons** element. Note that if you do not specify any ProgID in the Windows registry, Outlook will not call the **GetFormRegionManifest** and **GetFormRegionIcon** methods.


## See also


#### Concepts


[FormRegionStartup Interface](formregionstartup-object-outlook.md)

