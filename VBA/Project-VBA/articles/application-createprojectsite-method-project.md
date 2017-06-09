---
title: Application.CreateProjectSite Method (Project)
keywords: vbapj.chm142
f1_keywords:
- vbapj.chm142
ms.prod: project-server
ms.assetid: 79c77f3c-0ea6-eed7-762c-f364dc7f3ab7
ms.date: 06/08/2017
---


# Application.CreateProjectSite Method (Project)
Creates a site for the active project in a Project Web App instance on SharePoint Server 2013.

## Syntax

 _expression_. **CreateProjectSite** _(ParentSiteUrl,_ _NewSiteName,_ _LaunchBrowser)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ParentSiteUrl_|Optional|**Variant**|URL of the parent Project Web App site.|
| _NewSiteName_|Optional|**Variant**|Name of the new project site.|
| _LaunchBrowser_|Optional|**Variant**|**True** to launch the browser and open the new project site. The default value is **False**.|
| _ParentSiteUrl_|Optional|VARIANT||
| _NewSiteName_|Optional|VARIANT||
| _LaunchBrowser_|Optional|VARIANT||

### Return value

 **Boolean**

 **True** if the project site is successfully created.


