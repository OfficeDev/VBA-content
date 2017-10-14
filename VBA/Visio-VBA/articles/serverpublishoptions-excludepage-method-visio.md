---
title: ServerPublishOptions.ExcludePage Method (Visio)
keywords: vis_sdr.chm17962370
f1_keywords:
- vis_sdr.chm17962370
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.ExcludePage
ms.assetid: 3916ded4-daed-d6c7-9d75-c35273fed54a
ms.date: 06/08/2017
---


# ServerPublishOptions.ExcludePage Method (Visio)

Excludes the specified page from being published when the document is published as a VDW file.


## Syntax

 _expression_ . **ExcludePage**( **_PageNameU_** , **_Flags_** )

 _expression_ A variable that represents a **[ServerPublishOptions](serverpublishoptions-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PageName_|Required| **String**|The name of the page to exclude.|
| _Flags_|Required| **[VisLangFlags](vislangflags-enumeration-visio.md)**|Specifies whether  _PageName_ is local or universal.|

### Return Value

 **Nothing**


## Remarks

The  _Flags_ parameter must be one of the following **VisLangFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visLangLocal**|0|The page name is a local name.|
| **visLangUniversal**|1|The page name is a universal name.|

 **Note**  Excluding a page does not remove that page from the document?it merely prevents that page from appearing in the browser when the file is published as a VDW file. Because excluded pages remain in the document, they increase the size of the document and, hence, may negatively affect performance. For this reason, it is a good idea to use the  **[Page.Delete](page-delete-method-visio.md)** method to permanently delete unwanted pages from the document.

Calling the  **ExcludePage** method corresponds to clearing the check box for a page in the **Pages** list in the **Publish Settings** dialog box (click the **File** tab, click **Save &; Send**, click  **Save to SharePoint**, click  **Web Drawing (*.vdw)**, click  **Save As**, and then click  **Options**).


