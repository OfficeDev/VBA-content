---
title: Detect User Idle Time or Inactivity
ms.prod: access
ms.assetid: 40e9c4ef-a81b-074b-0be0-8247b4ea525b
ms.date: 06/08/2017
---


# Detect User Idle Time or Inactivity

This topic shows how to create a procedure that will run if your Access application does not detect any user input for a specified period of time. It involves creating a hidden form,  **DetectIdleTime**, which keeps track of idle time.

Follow these steps to create the  **DetectIdleTime** form.

1. Create a blank form that is not based on any table or query and name it  **DetectIdleTime**.

2. Set the following form properties:
    
     **Note**  The  **TimerInterval** setting indicates how often (in milliseconds) the application checks for user inactivity. A setting of 1000 equals 1 second.


  |**Property**|**Value**|
  |:-----|:-----|
  |OnTimer|[Event Procedure]|
  |TimerInterval|1000|

3. Enter the following code for the  **OnTimer** property event procedure:
    
```vb
Sub Form_Timer() 
         ' IDLEMINUTES determines how much idle time to wait for before 
         ' running the IdleTimeDetected subroutine. 
         Const IDLEMINUTES = 5 
 
         Static PrevControlName As String 
         Static PrevFormName As String 
         Static ExpiredTime 
 
         Dim ActiveFormName As String 
         Dim ActiveControlName As String 
         Dim ExpiredMinutes 
 
         On Error Resume Next 
 
         ' Get the active form and control name. 
 
         ActiveFormName = Screen.ActiveForm.Name 
         If Err Then 
            ActiveFormName = "No Active Form" 
            Err = 0 
         End If 
 
         ActiveControlName = Screen.ActiveControl.Name 
            If Err Then 
            ActiveControlName = "No Active Control" 
            Err = 0 
         End If 
 
         ' Record the current active names and reset ExpiredTime if: 
         '    1. They have not been recorded yet (code is running 
         '       for the first time). 
         '    2. The previous names are different than the current ones 
         '       (the user has done something different during the timer 
         '        interval). 
         If (PrevControlName = "") Or (PrevFormName = "") _ 
           Or (ActiveFormName <> PrevFormName) _ 
           Or (ActiveControlName <> PrevControlName) Then 
            PrevControlName = ActiveControlName 
            PrevFormName = ActiveFormName 
            ExpiredTime = 0 
         Else 
            ' ...otherwise the user was idle during the time interval, so 
            ' increment the total expired time. 
            ExpiredTime = ExpiredTime + Me.TimerInterval 
         End If 
 
         ' Does the total expired time exceed the IDLEMINUTES? 
         ExpiredMinutes = (ExpiredTime / 1000) / 60 
         If ExpiredMinutes >= IDLEMINUTES Then 
            ' ...if so, then reset the expired time to zero... 
            ExpiredTime = 0 
            ' ...and call the IdleTimeDetected subroutine. 
            IdleTimeDetected ExpiredMinutes 
         End If 
      End Sub
```

Then, create the following procedure in the form module:
    
```vb
Sub IdleTimeDetected(ExpiredMinutes) 
         Dim Msg As String 
         Msg = "No user activity detected in the last " 
         Msg = Msg &; ExpiredMinutes &; " minute(s)!" 
         MsgBox Msg, 48 
      End Sub
```

To hide the  **DetectIdleTime** form as it opens, set the _WindowMode_ argument of the **[OpenForm](docmd-openform-method-access.md)** method to **acHidden**.

