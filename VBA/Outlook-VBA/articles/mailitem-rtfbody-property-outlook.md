---
title: MailItem.RTFBody Property (Outlook)
keywords: vbaol11.chm3554
f1_keywords:
- vbaol11.chm3554
ms.prod: outlook
api_name:
- Outlook.MailItem.RTFBody
ms.assetid: 93bfda4f-08fb-9527-6946-625546d7fb49
ms.date: 06/08/2017
---


# MailItem.RTFBody Property (Outlook)

Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

 _expression_ . **RTFBody**

 _expression_ A variable that represents a **[MailItem](mailitem-object-outlook.md)** object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## Example

The following code samples in Microsoft Visual Basic for Applications (VBA) and C# displays the Rich Text Format body of the appointment in the active inspector. A  **MailItem** must be the active inspector for this code to work.


```vb
Sub GetRTFBodyForMail() 
 
 Dim oMail As Outlook.MailItem 
 
 Dim strRTF As String 
 
 If Application.ActiveInspector.CurrentItem.Class = olMail Then 
 
 Set oMail = Application.ActiveInspector.CurrentItem 
 
 strRTF = StrConv(oMail.RTFBody, vbUnicode) 
 
 Debug.Print strRTF 
 
 End If 
 
End Sub
```


```
private void GetRTFBodyForMail() 
 
{ 
 
 if (Application.ActiveInspector().CurrentItem is Outlook.MailItem) 
 
 { 
 
 Outlook.MailItem mail = 
 
 Application.ActiveInspector().CurrentItem as Outlook.MailItem; 
 
 byte[] byteArray = mail.RTFBody as byte[]; 
 
 System.Text.Encoding encoding = new System.Text.ASCIIEncoding(); 
 
 string RTF = encoding.GetString(byteArray); 
 
 Debug.WriteLine(RTF); 
 
 } 
 
} 
 

```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

