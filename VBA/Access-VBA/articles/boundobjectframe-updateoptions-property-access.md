---
title: BoundObjectFrame.UpdateOptions Property (Access)
keywords: vbaac10.chm10917
f1_keywords:
- vbaac10.chm10917
ms.prod: access
api_name:
- Access.BoundObjectFrame.UpdateOptions
ms.assetid: 919ad3b4-1128-947a-09c0-7c7b0373698e
ms.date: 06/08/2017
---


# BoundObjectFrame.UpdateOptions Property (Access)

You can use the  **UpdateOptions** property to specify how a linkedOLE object is updated. Read/write **Integer**.


## Syntax

 _expression_. **UpdateOptions**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **UpdateOptions** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Automatic|**acOLEUpdateAutomatic**|(Default) Updates the object each time the linked data changes.|
|Manual|**acOLEUpdateManual**|Updates the object only when the control's  **Action** property is set to **acOLEUpdate** or the link is updated with the **OLE/DDE Links** command on the **Edit** menu.|
Normally, the object is updated automatically whenever the linked data changes, but you can tell Microsoft Access to update the data only when it receives a specific instruction to do so. For example, if other users or applications can access or change linked spreadsheet data on a form, you can use this property to specify that the linked data only be updated when the database is opened in single-user mode.

When the  **UpdateOptions** property is set to Manual, updates don't occur based on the setting of the **Refresh interval** box on the **Advanced** tab of the **Options** dialog box, available by clicking **Options** on the **Tools** menu.


 **Note**  When an object's data is changed, the  **Updated** event occurs.


## Example

The following example sets the  **UpdateOptions** property for an unbound object frame named OLE1 to update manually, and then uses the **Action** property to force an update of the OLE object in the control.


```vb
OLE1.UpdateOptions = acOLEUpdateManual 
OLE1.Action = acOLEUpdate
```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

