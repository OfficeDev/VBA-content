---
title: Viewer.DisplayHelp Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.DisplayHelp
ms.assetid: 4d31b711-2521-cfd3-7689-0bd8618126b1
ms.date: 06/08/2017
---


# Viewer.DisplayHelp Method (Visio Viewer)

Displays the Help topic that has the specified ID in Microsoft Visio Viewer.


## Syntax

 _expression_. **DisplayHelp**( **_TopicID_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|TopicID|Required| **Long**|The ID of the Help topic to display.|

### Return Value

Nothing


## Remarks

The Help topic specified appears in the default browser.


## Example

The following code displays the default Help topic.


```
vsoViewer.DisplayHelp(1)
```


