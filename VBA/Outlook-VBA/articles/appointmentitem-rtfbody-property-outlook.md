---
title: AppointmentItem.RTFBody Property (Outlook)
keywords: vbaol11.chm3524
f1_keywords:
- vbaol11.chm3524
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.RTFBody
ms.assetid: 12af0270-e9bc-88ce-1d36-eafadf698406
ms.date: 06/08/2017
---


# AppointmentItem.RTFBody Property (Outlook)

Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

 _expression_ . **RTFBody**

 _expression_ A variable that represents an **[AppointmentItem](appointmentitem-object-outlook.md)** object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## Example

The following code samples in Microsoft Visual Basic for Applications (VBA) and C# displays the Rich Text Format body of the appointment in the active inspector. An  **AppointmentItem** must be the active inspector for this code to work.


```vb
Sub GetRTFBodyForMeeting() 
 
 Dim oAppt As Outlook.AppointmentItem 
 
 Dim strRTF As String 
 
 If Application.ActiveInspector.CurrentItem.Class = olAppointment Then 
 
 Set oAppt = Application.ActiveInspector.CurrentItem 
 
 strRTF = StrConv(oAppt.RTFBody, vbUnicode) 
 
 Debug.Print strRTF 
 
 End If 
 
End Sub
```


```
private void GetRTFBodyForAppt() 
 
{ 
 
 if (Application.ActiveInspector().CurrentItem is Outlook.AppointmentItem) 
 
 { 
 
 Outlook.AppointmentItem appt = 
 
 Application.ActiveInspector().CurrentItem as Outlook.AppointmentItem; 
 
 byte[] byteArray = appt.RTFBody as byte[]; 
 
 System.Text.Encoding encoding = new System.Text.ASCIIEncoding(); 
 
 string RTF = encoding.GetString(byteArray); 
 
 Debug.WriteLine(RTF); 
 
 } 
 
} 
 

```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

