---
title: Application.WindowActivate Method (Project)
keywords: vbapj.chm705
f1_keywords:
- vbapj.chm705
ms.prod: project-server
api_name:
- Project.Application.WindowActivate
ms.assetid: 8b9b39f8-39e5-b162-d8d9-de9838f7b39e
ms.date: 06/08/2017
---


# Application.WindowActivate Method (Project)

Activates a window.


## Syntax

 _expression_. **WindowActivate**( ** _WindowName_**, ** _DialogID_**, ** _TopPane_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _WindowName_|Optional|**String**|The name of the window to activate. The name of a window is the exact text that appears in the title bar of the window. The default is the name of the active window.|
| _DialogID_|Optional|**Long**|A constant specifying the dialog box to activate. Can be the following  **[PjDialog](pjdialog-enumeration-project.md)** constant: **pjResourceAssignment**.|
| _TopPane_|Optional|**Boolean**|**True** if Project should activate the upper pane. The default value is **True**.|

### Return Value

 **Boolean**


## Example

The following examples allow the user to specify and activate a "hot" window. If you assign the  **ActivateBookmarkedWindow** macro to a shortcut key, you can press that key to quickly activate the bookmarked window.


```vb
Public BookmarkedWindowName As String ' The name of the current bookmarked window 
 
Sub ActivateBookmarkedWindow() 
 
 Dim IsOpen As Boolean ' Whether or not the current bookmarked window is open 
 Dim I As Long ' Index for For...Next loop 
 
 IsOpen = False ' Assume the bookmarked window is not open. 
 
 For I = 1 To Windows.Count ' Look for the current bookmarked window. 
 If LCase(Windows(I).Caption) = LCase(BookmarkedWindowName) Then 
 IsOpen = True 
 Exit For 
 End If 
 Next I 
 
 ' If the current bookmarked window is not open or defined, then run 
 ' the ChangeBookmarkedWindow procedure. 
 If Len(BookmarkedWindowName) = 0 Or Not IsOpen Then 
 MsgBox ("The current bookmarked window is not open or has not been defined.") 
 ChangeBookmarkedWindowName 
 ' If the bookmarked window is open, activate it. 
 Else 
 WindowActivate (BookmarkedWindowName) 
 End If 
 
End Sub 
 
Sub ChangeBookmarkedWindowName() 
 
 Dim Entry As String ' The text entered by the user 
 
 Entry = InputBox$("Enter the name of the bookmarked window.") 
 
 ' If the user chooses Cancel, then exit the Sub procedure. 
 If Entry = Empty Then Exit Sub 
 
 ' Otherwise, set the name of the bookmarked window and then activate it. 
 BookmarkedWindowName = Entry 
 ActivateBookmarkedWindow 
 
End Sub
```


