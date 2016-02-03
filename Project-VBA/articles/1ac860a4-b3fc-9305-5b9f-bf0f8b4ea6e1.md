
# Application.MailRoutingSlip Method (Project)

Adds a mail routing slip for the active project.


## Syntax

 _expression_. **MailRoutingSlip**( **_To_**,  **_Subject_**,  **_Body_**,  **_AllAtOnce_**,  **_ReturnWhenDone_**,  **_TrackStatus_**,  **_Clear_**,  **_SendNow_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|To|Optional| **String**|The user names of the recipients of the message, separated by commas.|
|Subject|Optional| **String**| The subject of the message.|
|Body|Optional| **String**|The main text of the message.|
|AllAtOnce|Optional| **Boolean**| **True** if the message is sent to all users at the same time. **False** if the message is routed from one user to the next. The default value is **False**.|
|ReturnWhenDone|Optional| **Boolean**| **True** if the message returns to the sender after reaching the last recipient. The default value is **True**.|
|TrackStatus|Optional| **Boolean**| **True** if the location of the message is tracked. The default value is **True**.|
|Clear|Optional| **Boolean**| **True** if the list of user names in the **Routing Slip** dialog box is cleared. The default value is **False**.|
|SendNow|Optional| **Boolean**| **True** if the project is sent. **False** if the mail slip is edited without sending the project. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

Using the  **MailRoutingSlip** method without specifying any arguments displays the **Routing Slip** dialog box when a mail profile is set up on the user's system. If no mail profile is set up, using the **MailRoutingSlip** method without specifying any arguments displays the **Internet Connection Wizard**.


## Example

The following example sends the current schedule to Julie Rogers and then to Michael Edwards.


```vb
Sub PlanApproval() 
 MailRoutingSlip To:="Julie Rogers,Michael Edwards", _ 
 Subject:="Project Plan Approval", _ 
 Body:="Please review the following plan for approval.", _ 
 AllAtOnce:=False, ReturnWhenDone:=True, _ 
 TrackStatus:=True, SendNow:=True 
End Sub
```

