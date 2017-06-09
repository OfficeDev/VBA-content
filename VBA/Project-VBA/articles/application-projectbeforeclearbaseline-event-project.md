---
title: Application.ProjectBeforeClearBaseline Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeClearBaseline
ms.assetid: 4aa11658-7962-a46f-c914-5ed3bebd15a3
ms.date: 06/08/2017
---


# Application.ProjectBeforeClearBaseline Event (Project)

Occurs before a baseline is cleared. Uses the  **EventInfo** object parameter.


## Syntax

 _expression_. **ProjectBeforeClearBaseline**( ** _pj_**, ** _Interim_**, ** _bl_**, ** _InterimFrom_**, ** _AllTasks_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**| The project displayed in the deactivated window.|
| _Interim_|Required|**Boolean**|**True** if clearing an interim baseline plan. **False** if clearing a full baseline plan.|
| _bl_|Required|**PjBaselines**|The baseline you are clearing. Can be one of the following  **PjBaselines** constants: **pjBaseline**, **pjBaseline1**, **pjBaseline2**, **pjBaseline3**, **pjBaseline4**, **pjBaseline5**, **pjBaseline6**, **pjBaseline7**, **pjBaseline8**, **pjBaseline9**, or **pjBaseline10**.|
| _InterimFrom_|Required|**PjSaveBaselineTo**|The interim baseline plan being cleared. Can be one of the following  **PjSaveBaselineTo** constants: **pjIntoBaseline**, **pjIntoBaseline1**, **pjIntoBaseline2**, **pjIntoBaseline3**, **pjIntoBaseline4**, **pjIntoBaseline5**, **pjIntoBaseline6**, **pjIntoBaseline7**, **pjIntoBaseline8**, **pjIntoBaseline9**, **pjIntoBaseline10**, **pjIntoStart_Finish1**, **pjIntoStart_Finish2**, **pjIntoStart_Finish3**, **pjIntoStart_Finish4**, **pjIntoStart_Finish5**, **pjIntoStart_Finish6**, **pjIntoStart_Finish7**, **pjIntoStart_Finish8**, **pjIntoStart_Finish9**, or **pjIntoStart_Finish10**.|
| _AllTasks_|Required|**Boolean**|**True** if clearing the entire project.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the baseline is not cleared.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


## Example

The following sample displays a message box informing the user of a baseline clearing about to be made in the project plan. The message box indicates which baseline is being cleared (from 0 to 10), the file name of the project, and whether the interim plan is being cleared (True or False)




1. Create a new class module, and insert the following code:
    
```vb
Public WithEvents pApp As MSProject.Application 
Private Sub pApp_ProjectBeforeClearBaseline(ByVal pj As Project, _ 
 ByVal Interim As Boolean, ByVal bl As PjBaselines, _ 
 ByVal InterimFrom As PjSaveBaselineTo, _ 
 ByVal AllTasks As Boolean, ByVal Info As EventInfo) 
 
 MsgBox "Click OK to clear the baseline for the following " _ 
 &; "project:" &; vbCrLf &; "Baseline: " &; CStr(bl) _ 
 &; vbCrLf &; "Project: " &; pj.Name &; vbCrLf _ 
 &; "Clear interim plan: " &; CStr(Interim) 
End Sub
  ```


    
    
2. In a separate module, insert the following code:
    
```vb
Public X As New Class1 
Sub RunMacros() 
 Set X.pApp = MSProject.Application 
End Sub
  ```


    
    
3. Run the RunMacros procedure to start listening to the events.
    
4. On the  **Tools** menu, point to **Tracking** and click **Clear Baseline**.The event causes a message box to pop up every time a baseline is cleared.
    



