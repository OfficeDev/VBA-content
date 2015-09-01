
# TaskItem.Respond Method (Outlook)

 **Last modified:** July 28, 2015

Responds to a task request.

## Syntax

 _expression_. **Respond**( **_Response_**,  **_fNoUI_**,  **_fAdditionalTextDialog_**)

 _expression_A variable that represents a  **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Response|Required| ** [OlTaskResponse](7616cbdc-fc9c-abbe-fd07-ebdadc13ede2.md)**| The response to the request.|
|fNoUI|Required| **Variant**| **True** to not display a dialog box; the response is sent automatically. **False** to display the dialog box for responding.|
|fAdditionalTextDialog|Required| **Variant**| **False** to not prompt the user for input; the response is displayed in the inspector for editing. **True** to prompt the user to either send or send with comments. This argument is valid only iffNoUI is **False**.|

### Return Value

A  ** [TaskItem](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)** that represents the response to the task request.


## Remarks

When you call the  **Respond** method with the **olTaskAccept** parameter, Outlook creates a new **TaskItem** that duplicates the task request item. The new item has a different Entry ID. Outlook then removes the original item.

The following table describes the behavior of the  **Respond** method depending on the parent object, and thefNoUI andfAdditionalTextDialog parameters.



|**_fNoUI, fAdditionalTextDialog_**|**_Result_**|
|:-----|:-----|
| **True, True**|Response item is returned with no user interface. To send the response, you must call the  ** [Send](54f751fc-cff1-5d17-f635-f688cd8ad6f8.md)** method.|
| **True, False**|Same result as with  **True, True** .|
| **False, True**|If the  ** [Display](fea0619d-06dc-df44-fe93-5756eefb1be0.md)** method has been called, the user prompt appears. Otherwise, the item is sent without prompting and the resulting item is nothing.|
| **False, False**|Does nothing. |

## See also


#### Concepts


 [TaskItem Object](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)
#### Other resources


 [TaskItem Object Members](97234a76-2fc5-bbe4-2e14-25ae18694fc9.md)
