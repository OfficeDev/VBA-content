---
title: Application.OpenBrowser Method (Project)
keywords: vbapj.chm144
f1_keywords:
- vbapj.chm144
ms.prod: project-server
ms.assetid: 92691162-1c5f-43b6-57f2-8d56fa3f7bb6
ms.date: 06/08/2017
---


# Application.OpenBrowser Method (Project)
Opens the default web browser to a specified URL or the Windows Explorer to a specified directory or project file.

## Syntax

 _expression_. **OpenBrowser** _(URL)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _URL_|Optional|**String**|The URL to use for the browser address.|

### Return value

 **Boolean**

 **True** if the web browser or the Windows Explorer opens; otherwise, **False**.


## Remarks

You can use the  **OpenBrowser** method to open the browser to a specified URL. If the _URL_ parameter is not specified, the **OpenBrowser** method opens the Windows Explorer to the **My Documents** folder on the local computer.

If you specify an .MPP file path, Project opens the file.


## Examples

The following examples are valid, if the specified  _URL_ location exists:


-  `Application.OpenBrowser()`
    
-  `Application.OpenBrowser("http://MySharePointSite")`
    
-  `Application.OpenBrowser("http://MySharePointSite/_layouts/15/start.aspx#/Lists/Test%20tasks%20list%201/")`
    
-  `Application.OpenBrowser("file:///C:/Project")`
    
-  `Application.OpenBrowser("file://localhost/C|/Project")`
    
-  `Application.OpenBrowser("file:///C|/Project/Samples/Project1.mpp")`
    

## See also


#### Concepts


[Application Object](application-object-project.md)
