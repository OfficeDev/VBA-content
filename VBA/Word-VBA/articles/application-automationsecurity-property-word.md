---
title: Application.AutomationSecurity Property (Word)
keywords: vbawd10.chm158335425
f1_keywords:
- vbawd10.chm158335425
ms.prod: word
api_name:
- Word.Application.AutomationSecurity
ms.assetid: 2bc4f55c-d209-013b-77e4-ada7963bdee9
ms.date: 06/08/2017
---


# Application.AutomationSecurity Property (Word)

Returns or sets an  **MsoAutomationSecurity** constant that represents the security setting Microsoft Word uses when programmatically opening files. .


## Syntax

 _expression_ . **AutomationSecurity**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

The default setting of the  **AutomationSecurity** property is **msoAutomationSecurityLow** . Therefore, to avoid changing the users security settings or breaking solutions that rely on the default setting, you should be careful to set this property back to its original setting after programmatically opening a file.

Setting  **ScreenUpdating** to **False** does not affect alerts and will not affect security warnings. The **DisplayAlerts** setting will not apply to security warnings. For example, if the user sets **DisplayAlerts** equal to **False** and **AutomationSecurity** to **msoAutomationSecurityByUI** , while the user is on Medium security level, then there will be security warnings while a macro is running. This allows the macro to trap file open errors, while still showing the security warning if the file open succeeds.


## Example

This example changes the setting to disable macros, displays the  **Open** dialog box, and then sets the **AutomationSecurity** property back to its original setting.


```vb
Sub Security() 
 Dim lngAutomation As MsoAutomationSecurity 
 
 With Application 
 lngAutomation = .AutomationSecurity 
 .AutomationSecurity = msoAutomationSecurityForceDisable 
 With .FileDialog(msoFileDialogOpen) 
 .Show 
 .Execute 
 End With 
 .AutomationSecurity = lngAutomation 
 End With 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

