---
title: Application.GetCacheStatusForProject Property (Project)
ms.prod: project-server
ms.assetid: 71ab8ee0-83fc-c80f-3583-ce66b167d044
ms.date: 06/08/2017
---


# Application.GetCacheStatusForProject Property (Project)
Gets the state of a specified job that the active cache in Project Professional sends to the Project Server Queue System. Read-only  **PjCacheJobState**.

## Syntax

 _expression_. **GetCacheStatusForProject**

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProjectName_|Required|**String**|The name of the project; can be the active project or a different project that is open.|
| _ProjectJobType_|Required|**PjJobType**|Can be one of the [PjJobType](pjjobtype-enumeration-project.md) constants for the save, publish, or check-in operation.|

## Remarks

When you use Project Professional to perform an operation that uses one of the queue methods in Project Server, such as saving an update, publishing, or checking in a project, the Project Professional cache sends a job request to the Project Server Queue System. The  **GetCacheStatusForProject** property exposes the status of the queue job.


## Example

The  **TestCacheStatus** macro in the following example saves the active project, calls **WaitForJob** to wait for the queue to finish successfully, and then publishes the project. The **WaitForJob** macro periodically checks the job state by calling **GetCacheStatusForProject** and prints the job status to the **Immediate** window. If it finds the same status more than ten times in succession, the **WaitForJob** macro assumes there is a problem and exits. The example uses a **Sleep** method that can be run in either a 64-bit Project installation or a 32-bit Project installation.


```vb
Option Explicit

#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongLong)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Save and publish the active project; wait for the queue after each operation.
Sub TestCacheStatus()
    Const millisec2Wait = 500   ' Number of milliseconds to sleep between status messages.
    
    Application.FileSave
    If WaitForJob(PjJobType.pjCacheProjectSave, millisec2Wait) Then
        Debug.Print "Save completed ..."
    
        Application.Publish
        If WaitForJob(PjJobType.pjCacheProjectPublish, millisec2Wait) Then
            Debug.Print "Publish completed: " &; ActiveProject.Name
        End If
    Else
        Debug.Print "Save job not completed"
    End If
End Sub

' Check the cache job state for a save, publish, or check-in operation.
Function WaitForJob(job As PjJobType, msWait As Long) As Boolean
    ' Number of times the same job status is repeated until WaitForJob exits with error.
    Const repeatedLimit = 10
    
    Dim jobState As Integer
    Dim previousJobState As Integer
    Dim bail As Integer
    Dim jobType As String
    
#If Win64 Then
    Dim millisec As LongLong
    millisec = CLngLng(msWait)
#Else
    Dim millisec As Long
    millisec = msWait
#End If

    WaitForJob = True
    
    Select Case job
        Case PjJobType.pjCacheProjectSave
            jobType = "Save"
        Case PjJobType.pjCacheProjectPublish
            jobType = "Publish"
        Case PjJobType.pjCacheProjectCheckin
            jobType = "Checkin"
        Case Else
            jobType = "unknown"
    End Select

    bail = 0
    
    If (jobType = "unknown") Then
        WaitForJob = False
    Else
        Do
            jobState = Application.GetCacheStatusForProject(ActiveProject.Name, job)
            Debug.Print jobType &; " job state: " &; jobState
            
            ' Bail out if something is wrong.
            If jobState = previousJobState Then bail = bail + 1
            If bail > repeatedLimit Then
                WaitForJob = False
                Exit Do
            End If
            
            previousJobState = jobState
            
            Sleep (msWait)
        Loop While Not (jobState = PjCacheJobState.pjCacheJobStateSuccess)
    End If
End Function
```

Following is the output for a wait time of 500 milliseconds between status messages. If the network latency is greater, set the wait time for a longer interval. To find the meaning of output values, see the [PjCacheJobState](pjcachejobstate-enumeration-project.md) enumeration. For example, the value **4** is the **pjCacheJobStateSuccess** constant. If you run **TestCacheStatus** when there are no changes made to the project, the save job state repeats many times as **-1**, which is the value of the  **pjCacheJobStateInvalid** constant.




```
Save job state: 4
Save completed ...
Publish job state: -1
Publish job state: 3
Publish job state: 3
Publish job state: 4
Publish completed: WinProj test 1
```


## Property value

 **PJCACHEJOBSTATE**


## See also


#### Other resources


[PjCacheJobState Enumeration](pjcachejobstate-enumeration-project.md)
[PjJobType Enumeration](pjjobtype-enumeration-project.md)
