---
title: Application.WindowBeforeDoubleClick Event (Word)
keywords: vbawd10.chm4000013
f1_keywords:
- vbawd10.chm4000013
ms.prod: word
api_name:
- Word.Application.WindowBeforeDoubleClick
ms.assetid: ece03591-0410-9dac-dedf-72c736dd477a
ms.date: 06/08/2017
---


# Application.WindowBeforeDoubleClick Event (Word)

Occurs when the editing area of a document window is double-clicked, before the default double-click action.


## Syntax

 _expression_ . **Private Sub object_WindowBeforeDoubleClick**( **_ByVal Sel As Selection_** , **_Cancel As Boolean_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For more information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sel_|Required| **Selection**|The current selection.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the default double-click action does not occur when the procedure is finished.|

## Example

This example prompts the user for a yes or no response before executing the default double-click action. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_WindowBeforeDoubleClick _ 
 (ByVal Sel As Selection, Cancel As Boolean) 
 Dim intResponse As Integer 
 intResponse = MsgBox("Selection = " &; Sel &; vbLf &; vbLf _ 
 &; "Continue with operation on this selection?", _ 
 vbYesNo) 
 If intResponse = vbNo Then Cancel = True 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

