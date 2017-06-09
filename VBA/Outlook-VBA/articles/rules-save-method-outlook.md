---
title: Rules.Save Method (Outlook)
keywords: vbaol11.chm2161
f1_keywords:
- vbaol11.chm2161
ms.prod: outlook
api_name:
- Outlook.Rules.Save
ms.assetid: d838eca0-4ec5-ab43-a031-fd65ab7d9f3c
ms.date: 06/08/2017
---


# Rules.Save Method (Outlook)

Saves all rules in the  **[Rules](rules-object-outlook.md)** collection.


## Syntax

 _expression_ . **Save**( **_ShowProgress_** )

 _expression_ A variable that represents a **Rules** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShowProgress_|Optional| **Boolean**| **True** to display the progress dialog box, **False** to save rules without showing the progress.|

## Remarks

After you enable a rule, you must also save the rule by using  **Rules.Save** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.

 **Rules.Save** can be an expensive operation in terms of performance on slow connections to Exchange server. For more information on using the progress dialog box, see[Manage Rules in the Outlook Object Model](http://msdn.microsoft.com/library/05ddd643-e9bd-a37d-b680-b8519960a5f6%28Office.15%29.aspx).

Saving rules that are incompatible or have improperly defined actions or conditions (such as an empty string for  **[TextRuleCondition.Text](textrulecondition-text-property-outlook.md)** ) will return an error.

The Exchange server limits the maximum number of rules that can be supported by a store.  **Rules.Save** returns an error when this limit is reached.


## See also


#### Concepts


[Rules Object](rules-object-outlook.md)

