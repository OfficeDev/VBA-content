---
title: Rule.Execute Method (Outlook)
keywords: vbaol11.chm2173
f1_keywords:
- vbaol11.chm2173
ms.prod: outlook
api_name:
- Outlook.Rule.Execute
ms.assetid: 487abb6f-9003-04a4-f4e2-3f66b3ba5a52
ms.date: 06/08/2017
---


# Rule.Execute Method (Outlook)

Applies a rule as an one-off operation.


## Syntax

 _expression_ . **Execute**( **_ShowProgress_** , **_Folder_** , **_IncludeSubfolders_** , **_RuleExecuteOption_** )

 _expression_ A variable that represents a **Rule** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShowProgress_|Optional| **Boolean**| **True** to display the progress dialog box when the rule is executed, **False** to run the rule without displaying the dialog box.|
| _Folder_|Optional| **[Folder](folder-object-outlook.md)**|Represents the folder where the rule will be applied.|
| _IncludeSubfolders_|Optional| **Boolean**| **True** to apply the rule to subfolders of the folder indicated by the _Folder_ parameter; **False** to apply the rule only to that folder but not its subfolders.|
| _RuleExecuteOption_|Optional| **[OlRuleExecuteOption](olruleexecuteoption-enumeration-outlook.md)**|Represents whether to apply the rule to read, unread, or all messages in the folder or folders specified by the  _Folder_ and _IncludeSubfolders_ parameters.|

## Remarks

Use  **[Rule.Execute](rule-execute-method-outlook.md)** to apply a rule as a one-off operation regardless of whether **[Rule.Enabled](rule-enabled-property-outlook.md)** is **True** . Use **Rule.Enabled** and then **[Rules.Save](rules-save-method-outlook.md)** if you want to apply the rule consistently and persist the rules beyond the current session.

The parameters to the  **Execute** method are optional. If you do not specify any parameters, the rule will be applied to all messages in the Inbox but not to the subfolders of the Inbox. The default values for the optional arguments are as follows:



| **Parameter**| **Default Value**|
| _ShowProgress_| **False**|
| _Folder_|Inbox|
| _IncludeSubfolders_| **False**|
| _RuleExecuteOption_| **OlRuleExecuteOption.olRuleExecuteAllMessages**|
If  _ShowProgress_ is **True** and the user cancels the progress dialog box, rule execution is canceled in the same manner as if the user had canceled rule execution through the **Rules and Alerts Wizard**.  **Execute** returns an error when the user cancels the progress dialog.

If you plan to show a custom progress user interface instead of using the progress dialog box, you should be aware that there are no events that indicate when rule execution starts and stops. 


## See also


#### Concepts


[Rule Object](rule-object-outlook.md)

