---
title: IAssistance.SetDefaultContext Method (Office)
ms.prod: office
api_name:
- Office.IAssistance.SetDefaultContext
ms.assetid: 3eea8f7a-12a3-aca4-f963-28c5c4e63c96
ms.date: 06/08/2017
---


# IAssistance.SetDefaultContext Method (Office)

Sets a help topic as the default topic that will be displayed when the user opens a help window.


## Syntax

 _expression_. **SetDefaultContext**( **_HelpId_** )

 _expression_ An expression that returns a **IAssistance** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HelpId_|Required|**String**|The ID of the default help topic.|

## Remarks

The help topic specified in this method will not be displayed if Office has already defined a default help topic in that scope. In addition, for some dialog boxes, regardless of whether an ID is passed by the method, the help topic that shipped with Office will be displayed when the user presses the  **F1** key or clicks the **Help** button. For example, if the user is in a custom dialog box and presses **F1**, a custom or built-in help topic will be displayed if one is specified by the developer. If no default ID is specified, the default built-in Office help topic will be displayed. Likewise, if the user is in the  **New Document** dialog box, for example, the Office specified help topic will be displayed regardless of whether a different ID is passed by the method.

The  **Assistance** property returns an **IAssistance** object. The **IAssistance** object exposes methods that allow developers to display help topics in the Office Help Viewer or to display help topics that ship with Office in the Help window of the host application. Developers either pass specific Help IDs to the help system or pass specific search queries. Help IDs have to be explicitly added to the Help file in order for the Help ID to return the help topic.


## Example

The following example, the help topic associated with ID "60385" will be set as the default for the application. 


```
Sub SetDefaultHelpTopic() 
 Application.Assistance.SetDefaultContext "60385" 
End Sub
```


## See also


#### Concepts


[IAssistance Object](iassistance-object-office.md)
#### Other resources


[IAssistance Object Members](iassistance-members-office.md)

