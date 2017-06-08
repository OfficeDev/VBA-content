---
title: Application.ProtectedViewWindowDeactivate Event (Word)
keywords: vbawd10.chm4000035
f1_keywords:
- vbawd10.chm4000035
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindowDeactivate
ms.assetid: bd80056b-edce-7e0b-c61a-31ebda24a416
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowDeactivate Event (Word)

Occurs when a protected view window is deactivated.


## Syntax

 _expression_ . **ProtectedViewWindowDeactivate**( **_PvWindow_** , )

 _expression_ An expression that returns a **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PvWindow_|Required| **[ProtectedViewWindow](protectedviewwindow-object-word.md)**|The deactivated protected view window.|

## Example

The following code example minimizes an open protected view window when it is deactivated. This code must be placed in a class module, and an instance of the class must be correctly initialized for this code example to work correctly. For more information about how to do this, see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```vb
Private Sub App_ProtectedViewWindowDeactivate(ByVal PvWindow As ProtectedViewWindow) 
 PvWindow.WindowState = wdWindowStateMinimize 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

