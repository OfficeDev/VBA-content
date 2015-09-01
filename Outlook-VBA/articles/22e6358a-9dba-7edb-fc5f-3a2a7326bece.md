
# Inspector.SetSchedulingStartTime Method (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets the start time for a meeting item in the free/busy grid on the  **Scheduling Assistant** tab of the inspector.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SetSchedulingStartTime**( **_Start_**)

 _expression_A variable that represents an  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Start|Required| **Date**|The beginning of the time range that the  **Scheduling Assistant** tab of the inspector displays free/busy times for meeting attendees.|

## Remarks
<a name="sectionSection1"> </a>

The object specified by the  ** [CurrentItem](eaaf0192-a169-c107-95a6-b8e759a3b873.md)** property of the parent ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)** object must be an ** [AppointmentItem](204a409d-654e-27aa-643a-8344c631b82d.md)** or ** [MeetingItem](b75730f5-b395-3d66-5acd-b64fd8fcd78f.md)**. The  **Scheduling Assistant** tab must be displayed in the inspector, otherwise Microsoft Outlook raises an error. If Outlook cannot display the **Scheduling Assistant** tab for that item type, Outlook displays the following error: **The scheduling start time can only be set when the Scheduling Assistant is displayed on a meeting item.**


## Example
<a name="sectionSection2"> </a>

The following code sample in Microsoft Visual Basic for Applications (VBA) shows how to use the  **SetSchedulingStartTime** method to set the scheduling start time on the **Scheduling Assistant** tab of an **AppointmentItem**. The appointment start time is set to one month from now, and the scheduling start time is also set to one month from now.


```
Sub DemoSetSchedulingStartTime() 
 
 Dim oAppt As Outlook.AppointmentItem 
 
 Dim oInsp As Outlook.inspector 
 
 
 
 ' Create and display appointment. 
 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 
 oAppt.MeetingStatus = olMeeting 
 
 oAppt.Subject = "Test Appointment" 
 
 oAppt.Start = DateAdd("m", 1, Now) 
 
 ' Display the appointment in the Appointment tab of the inspector. 
 
 oAppt.Display 
 
 
 
 Set oInsp = oAppt.GetInspector 
 
 ' Switch to the Scheduling Assistant tab in that inspector. 
 
 oInsp.SetCurrentFormPage ("Scheduling Assistant") 
 
 ' Set the appointment start time in the Scheduling Assistant. 
 
 oInsp.SetSchedulingStartTime (DateAdd("m", 1, Now)) 
 
End Sub 
 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Inspector Object](d7384756-669c-0549-1032-c3b864187994.md)
#### Other resources


 [Inspector Object Members](acd3e13f-4727-7966-d2a5-a95e4528425c.md)
