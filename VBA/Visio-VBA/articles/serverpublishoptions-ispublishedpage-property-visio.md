---
title: ServerPublishOptions.IsPublishedPage Property (Visio)
keywords: vis_sdr.chm17962635
f1_keywords:
- vis_sdr.chm17962635
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.IsPublishedPage
ms.assetid: b174f50d-4d37-962a-06cc-5013b36309ff
ms.date: 06/08/2017
---


# ServerPublishOptions.IsPublishedPage Property (Visio)

Returns  **True** if the specified page is designated to be included when the document is published as a .vdw file. Read-only.


## Syntax

 _expression_ . **IsPublishedPage**( **_PageName_** **_Flags_** )

 _expression_ A variable that represents a **[ServerPublishOptions](serverpublishoptions-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PageName_|Required| **String**|The page to check for publish status.|
| _Flags_|Required| **[VisLangFlags](vislangflags-enumeration-visio.md)**|Specifies whether the page name is a local or a universal name.|

### Return Value

 **Boolean**


## Remarks

The setting of the  **IsPublishedPage** property corresponds to the status (selected or cleared) of the box that represents the specified page in the **Pages** list in the **Publish Settings** dialog box. (Click the **File** tab, click **Save &; Send**, click  **Save to SharePoint**, click  **Web Drawing (*.vdw)**, click  **Save As**, and then click  **Options**.) The default is for all pages in the document to be designated for publishing.

To change the publish status of a page, you can use the  **[IncludePage](serverpublishoptions-includepage-method-visio.md)** and **[ExcludePage](serverpublishoptions-excludepage-method-visio.md)** methods of the **ServerPublishOptions** object.


