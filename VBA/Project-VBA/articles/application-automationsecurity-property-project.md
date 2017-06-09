---
title: Application.AutomationSecurity Property (Project)
ms.prod: project-server
api_name:
- Project.Application.AutomationSecurity
ms.assetid: 08f71d7f-37bf-c845-89c3-a69e34892efe
ms.date: 06/08/2017
---


# Application.AutomationSecurity Property (Project)

Gets or sets a value that represents the security mode that Project uses when programmatically opening files. Read/write  **MsoAutomationSecurity**.


## Syntax

 _expression_. **AutomationSecurity**

 _expression_ A variable that represents an **Application** object.


## Remarks

The default value of the  **AutomationSecurity** property is **msoAutomationSecurityByUI**. The value can be one of the following **MsoAutomationSecurity** constants:



|**Constant**|**Description**|
|:-----|:-----|
|**msoAutomationSecurityByUI**|Uses the security setting specified on the  **Macro Settings** tab of the **Trust Center** dialog box.|
|**msoAutomationSecurityForceDisable**|Disables all macros in all files opened programmatically without showing any security alerts.|
|**msoAutomationSecurityLow**|Enables all macros. This value is not recommended because potentially dangerous code can run.|
 **Macro Settings** tab of the **Trust Center** dialog box has four settings for the macro security level. The default setting is **Disable all macros with notification**. For more information about security settings and digital code signing, see the links on the  **Trust Center** tab of the **Project Options** dialog box.


