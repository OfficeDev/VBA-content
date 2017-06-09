---
title: Application.ProtectedViewWindowActivate Event (Word)
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindowActivate
ms.assetid: ae68e1aa-7cec-cd76-ee0e-71a051c5b6e3
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowActivate Event (Word)

Occurs when any protected view window is activated.


## Syntax

 _expression_ . **ProtectedViewWindowActivate**( **_PvWindow_** , )

 _expression_ An expression that returns a **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PvWindow_|Required| **[ProtectedViewWindow](protectedviewwindow-object-word.md)**|The protected view window that is activated.|

## Example

The following code example maximizes any protected view window when it is activated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work. For more information about how to do this, see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```vb
Private Sub App_ProtectedViewWindowActivate(ByVal PvWindow As ProtectedViewWindow) 
 PvWindow.WindowState = wdWindowStateMaximize 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

