---
title: Application.LoadWebPaneControl Method (Project)
keywords: vbapj.chm56
f1_keywords:
- vbapj.chm56
ms.prod: project-server
api_name:
- Project.Application.LoadWebPaneControl
ms.assetid: b807a6e0-5a85-14a0-a87f-e4b6181c9648
ms.date: 06/08/2017
---


# Application.LoadWebPaneControl Method (Project)

Supports the Web pane that hosts the  **Task Drivers**,  **Project/Resource Import Wizard**, and  **Deliverables** features.


## Syntax

 _expression_. **LoadWebPaneControl**( ** _TargetPage_**, ** _WrapperPage_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TargetPage_|Required|**String**|A numeric ID that identifies the HTML target page that needs to be displayed.  **TargetPage** can also be set to a URL, an XML stream, a pointer to an XML file, or any other string value.|
| _WrapperPage_|Optional|**Variant**|A pointer to an HTML page that provides wrapper functionality for the page being displayed in Project. The wrapper page contains event-handling code that allows Project functionality, such as saving files or changing views, to work when a web page is being displayed. The WrapperPage parameter is only used when the  **Project Guid**e is hidden. When the  **Project Guide** is shown, mainpage.htm is used as the wrapper page, and a **WrapperPage** parameter, if specified, is ignored. If no WrapperPage parameter is specified, Project's default wrapper page, gbui://wrapper.htm, is used|

### Return Value

 **Boolean**


## Remarks

 The **LoadWebPaneControl** method is similar to the **LoadWebBrowserControl** method for the **Project Guide**, except TargetPage is a URL and the method generates an  **Application.LoadWebPane** event. The default WrapperPage is Mainpage_wp.htm.


