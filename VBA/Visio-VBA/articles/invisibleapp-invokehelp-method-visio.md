---
title: InvisibleApp.InvokeHelp Method (Visio)
keywords: vis_sdr.chm17550695
f1_keywords:
- vis_sdr.chm17550695
ms.prod: visio
api_name:
- Visio.InvisibleApp.InvokeHelp
ms.assetid: e3860d89-8d07-22d8-664b-b12becd39d98
ms.date: 06/08/2017
---


# InvisibleApp.InvokeHelp Method (Visio)

Performs operations that involve the Microsoft Visio Help system.


## Syntax

 _expression_ . **InvokeHelp**( **_bstrHelpFileName_** , **_Command_** , **_Data_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrHelpFileName_|Required| **String**|Specifies an HTML file, a URL, a compiled HTML file, or an optional window definition (preceded with a ">" character). If the command being used does not require a file or URL, this value may be "".|
| _Command_|Required| **Long**|The action to perform.|
| _Data_|Required| **Long**|Any data that is required based on the value of the command argument.|

### Return Value

Nothing


## Remarks

Using the  **InvokeHelp** method, you can create a custom Help system that is integrated with the Visio Help system. To enable your custom Help to appear in the same MSO Help window as Visio Help, do not specify a window definition in the _bstrHelpFileName_ argument.

The arguments passed to the  **InvokeHelp** method correspond to those described in the HTML Help API. For a list of command values, see the HTML Help API Reference on MSDN, the Microsoft Developer Network. Microsoft Visual Basic programmers can use the numeric equivalent of the C++ constants defined in the HTML Help API header files.

For example, use the following code to display the Visio Help window:




```vb
Application.InvokeHelp "Visio.chm", 15, 0
```

For more information about the HTML Help API, search for "HTML Help API overview" on MSDN.


