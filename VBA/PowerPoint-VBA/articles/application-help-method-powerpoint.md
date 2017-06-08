---
title: Application.Help Method (PowerPoint)
keywords: vbapp10.chm502021
f1_keywords:
- vbapp10.chm502021
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Help
ms.assetid: 97dabc76-1987-6e08-ea42-6762be6b7d60
ms.date: 06/08/2017
---


# Application.Help Method (PowerPoint)

Displays a Help topic.


## Syntax

 _expression_. **Help**( **_HelpFile_**, **_ContextID_** )

 _expression_ A variable that represents a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HelpFile_|Optional|**String**|The name of the Help file you want to display. Can be either a .chm or an .hlp file. If this argument is not specified, Microsoft PowerPoint Help is used.|
| _ContextID_|Optional|**Long**|The context ID number for the Help topic. If this argument is not specified or if it specifies a context ID number that is not associated with a Help topic, the  **Help Topics** dialog box is displayed.|

## Example

This example displays topic number 65527 in the Help file MyHelpFile.chm.


```vb
Application.Help "MyHelpFile.chm", 65527
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

