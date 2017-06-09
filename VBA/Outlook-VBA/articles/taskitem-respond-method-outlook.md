---
title: TaskItem.Respond Method (Outlook)
keywords: vbaol11.chm1754
f1_keywords:
- vbaol11.chm1754
ms.prod: outlook
api_name:
- Outlook.TaskItem.Respond
ms.assetid: 1befabf7-262f-897a-d1dc-49be4e7ddf9b
ms.date: 06/08/2017
---


# TaskItem.Respond Method (Outlook)

Responds to a task request.


## Syntax

 _expression_ . **Respond**( **_Response_** , **_fNoUI_** , **_fAdditionalTextDialog_** )

 _expression_ A variable that represents a **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **[OlTaskResponse](oltaskresponse-enumeration-outlook.md)**| The response to the request.|
| _fNoUI_|Required| **Variant**| **True** to not display a dialog box; the response is sent automatically. **False** to display the dialog box for responding.|
| _fAdditionalTextDialog_|Required| **Variant**| **False** to not prompt the user for input; the response is displayed in the inspector for editing. **True** to prompt the user to either send or send with comments. This argument is valid only if _fNoUI_ is **False** .|

### Return Value

A  **[TaskItem](taskitem-object-outlook.md)** that represents the response to the task request.


## Remarks

When you call the  **Respond** method with the **olTaskAccept** parameter, Outlook creates a new **TaskItem** that duplicates the task request item. The new item has a different Entry ID. Outlook then removes the original item.

The following table describes the behavior of the  **Respond** method depending on the parent object, and the _fNoUI_ and _fAdditionalTextDialog_ parameters.



|**_fNoUI, fAdditionalTextDialog_**|**_Result_**|
|:-----|:-----|
| **True, True**|Response item is returned with no user interface. To send the response, you must call the  **[Send](taskitem-send-method-outlook.md)** method.|
| **True, False**|Same result as with  **True, True** .|
| **False, True**|If the  **[Display](taskitem-display-method-outlook.md)** method has been called, the user prompt appears. Otherwise, the item is sent without prompting and the resulting item is nothing.|
| **False, False**|Does nothing. |

## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

