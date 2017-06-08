---
title: BoundObjectFrame.OLEType Property (Access)
keywords: vbaac10.chm10958
f1_keywords:
- vbaac10.chm10958
ms.prod: access
api_name:
- Access.BoundObjectFrame.OLEType
ms.assetid: 9ce7cb88-e13e-4cda-bfe7-096734b796a0
ms.date: 06/08/2017
---


# BoundObjectFrame.OLEType Property (Access)

You can use the  **OLEType** property to determine if a control contains an OLE object, and, if so, whether the object is linked or embedded. Read/write **Byte**.


## Syntax

 _expression_. **OLEType**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **OLEType** property uses the following settings.



|**Setting**|**Constant**|**Description**|
|:-----|:-----|:-----|
|Linked|**acOLELinked**|The control contains a linked object. All the object's data is managed by the application that created it.|
|Embedded|**acOLEEmbedded**|The control contains an embedded object. All the object's data is managed by Microsoft Access.|
|None|**acOLENone**|The control doesn't contain an OLE object.|
When creating an OLE object, use the  **OLETypeAllowed** property to determine what type of object a control can contain.


## Example

The following example illustrates how to display the  **Insert Object** dialog box and how to display an error message if the **Cancel** button in the **Insert Object** dialog box is clicked.


```vb
Sub InsertObject_Click() 
 Dim conUserCancelled As Integer 
 
 ' Error message returned when user cancels. 
 conUserCancelled = 2001 
 On Error GoTo ButtonErr 
 If OLE1.OLEType = acOLENone Then 
 ' No OLE object created. 
 ' Display Insert Object dialog box. 
 OLE1.Action = acOLEInsertObjDlg 
 End If 
 Exit Sub 
 
ButtonErr: 
 If Err = conUserCancelled Then ' Display message. 
 MsgBox "You clicked the Cancel button; " _ 
 &; "no object was created." 
 End If 
 Resume Next 
End Sub
```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

