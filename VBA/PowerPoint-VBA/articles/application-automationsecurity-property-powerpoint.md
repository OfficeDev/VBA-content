---
title: Application.AutomationSecurity Property (PowerPoint)
keywords: vbapp10.chm502048
f1_keywords:
- vbapp10.chm502048
ms.prod: powerpoint
api_name:
- PowerPoint.Application.AutomationSecurity
ms.assetid: 942341fe-5290-2903-db70-4e7cff0d75c7
ms.date: 06/08/2017
---


# Application.AutomationSecurity Property (PowerPoint)

Represents the security mode that Microsoft PowerPoint uses when it opens files programmatically. Read/write.


## Syntax

 _expression_. **AutomationSecurity**

 _expression_ A variable that represents an **Application** object.


### Return Value

MsoAutomationSecurity


## Remarks

This property is automatically set to  **msoAutomationSecurityLow** when the application is started. Therefore, to avoid breaking solutions that rely on the default setting, you should be careful to reset this property to **msoAutomationSecurityLow** after programmatically opening a file. Also, to avoid malicious subversion, you should set this property immediately before and after you open a file programmatically .

The value of the  **[DisplayAlerts](application-displayalerts-property-powerpoint.md)** property does not apply to security warnings. For example, if the user sets the **DisplayAlerts** property equal to **False** and the **AutomationSecurity** property to **msoAutomationSecurityByUI**, while the user is on Medium security level, there will be security warnings while the macro is running. This allows the macro to trap file open errors, while still showing the security warning if the file succeeds in opening.

The value of the  **AutomationSecurity** property can be one of these **MsoAutomationSecurity** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoAutomationSecurityByUI**|Uses the security setting specified in the  **Trust Center** dialog box.|
|**msoAutomationSecurityForceDisable**| Disables all macros in all files opened programmatically without showing any security alerts.|
|**msoAutomationSecurityLow**|Enables all macros. This is the default value when the application is started.|

## Example

This example captures the current automation security setting, changes the setting to disable macros, displays the  **Open** dialog box, and after opening the selected presentation, sets the automation security back to its original setting.


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


[Application Object](application-object-powerpoint.md)

