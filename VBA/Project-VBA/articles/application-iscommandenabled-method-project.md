---
title: Application.IsCommandEnabled Method (Project)
keywords: vbapj.chm131102
f1_keywords:
- vbapj.chm131102
ms.prod: project-server
api_name:
- Project.Application.IsCommandEnabled
ms.assetid: 22202fed-7531-0f87-0e38-3ee703717ec1
ms.date: 06/08/2017
---


# Application.IsCommandEnabled Method (Project)

Shows whether the specified command is enabled.


## Syntax

 _expression_. **IsCommandEnabled**( ** _CommandName_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CommandName_|Required|**String**|The name of a valid command.|

### Return Value

 **Long**


## Remarks

Valid commands are VBA method names in the  **MSProject** library. The return value can be one of the **[PjIsCommandEnabled](pjiscommandenabled-enumeration-project.md)** constants.


## Example

When the Team Planner view is not visible, the  **TestCommandEnabled** macro returns the following results:



The  **FileOpen** method is available in most cases. The **IsCommandEnabled** method is undefined because it is not included in the internal list of methods. The **ResetTPStyle** method is disabled because it is only available when the Team Planner view is open.




```vb
Sub TestCommandEnabled() 
 Dim commandArray(3) As String 
 Dim isEnabled As String 
 Dim i As Integer 
 
 commandArray(1) = "FileOpen" 
 commandArray(2) = "IsCommandEnabled" 
 commandArray(3) = "ResetTPStyle" 
 
 For i = 1 To 3 
 isEnabled = GetCommandEnabled(commandArray(i)) 
 Debug.Print commandArray(i) &; " is " &; isEnabled 
 Next i 
End Sub 
 
Function GetCommandEnabled(command As String) As String 
 Dim isEnabled As Long 
 Dim enabledMsg As String 
 Dim result As String 
 
 isEnabled = Application.IsCommandEnabled(command) 
 
 Select Case isEnabled 
 Case PjIsCommandEnabled.pjCommandDisabled 
 result = "disabled." 
 Case PjIsCommandEnabled.pjCommandEnabled 
 result = "enabled." 
 Case PjIsCommandEnabled.pjCommandUndefined 
 result = "undefined." 
 Case Else 
 result = "unknown result." 
 End Select 
 
 GetCommandEnabled = result 
End Function
```


