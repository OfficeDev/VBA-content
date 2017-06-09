---
title: Application.AutomationSecurity Property (Excel)
keywords: vbaxl10.chm133269
f1_keywords:
- vbaxl10.chm133269
ms.prod: excel
api_name:
- Excel.Application.AutomationSecurity
ms.assetid: ae19bf93-dc0f-f18a-d8ce-f54108602844
ms.date: 06/08/2017
---


# Application.AutomationSecurity Property (Excel)

Returns or sets an  **[MsoAutomationSecurity](http://msdn.microsoft.com/library/6147cad7-3db3-7f9a-397e-62dd64b89b50%28Office.15%29.aspx)** constant that represents the security mode Microsoft Excel uses when programmatically opening files. Read/write.


## Syntax

 _expression_ . **AutomationSecurity**

 _expression_ A variable that represents an **Application** object.


## Remarks

This property is automatically set to  **msoAutomationSecurityLow** when the application is started. Therefore, to avoid breaking solutions that rely on the default setting, you should be careful to reset this property to **msoAutomationSecurityLow** after programmatically opening a file. Also, this property should be set immediately before and after opening a file programmatically to avoid malicious subversion.



**MsoAutomationSecurity** can be one of these **MsoAutomationSecurity** constants.
- **msoAutomationSecurityByUI** . Uses the security setting specified in the **Security** dialog box.|
- **msoAutomationSecurityForceDisable** . Disables all macros in all files opened programmatically without showing any security alerts.<table><tr><th>**Note**</th></tr><tr><td>This setting does not disable Microsoft Excel 4.0 macros. If a file that contains Microsoft Excel 4.0 macros is opened programmatically, the user will be prompted to decide whether or not to open the file.</td></tr></table>
- **msoAutomationSecurityLow** . Enables all macros. This is the default value when the application is started.

Setting  **[ScreenUpdating](application-screenupdating-property-excel.md)** to **False** does not affect alerts and will not affect security warnings. The **[DisplayAlerts](application-displayalerts-property-excel.md)** setting will not apply to security warnings. For example, if the user sets **DisplayAlerts** equal to **False** and **AutomationSecurity** to **msoAutomationSecurityByUI** , while the user is on Medium security level, then there will be security warnings while the macro is running. This allows the macro to trap file open errors, while still showing the security warning if the file open succeeds.


## Example

This example captures the current automation security setting, changes the setting to disable macros, displays the  **Open** dialog box, and after opening the selected document, sets the automation security back to its original setting.


```vb
Sub Security() 
    Dim secAutomation As MsoAutomationSecurity 
 
    secAutomation = Application.AutomationSecurity 
 
    Application.AutomationSecurity = msoAutomationSecurityForceDisable 
    Application.FileDialog(msoFileDialogOpen).Show 
 
    Application.AutomationSecurity = secAutomation 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

