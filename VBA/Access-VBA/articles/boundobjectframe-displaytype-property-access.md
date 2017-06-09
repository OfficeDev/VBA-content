---
title: BoundObjectFrame.DisplayType Property (Access)
keywords: vbaac10.chm10916
f1_keywords:
- vbaac10.chm10916
ms.prod: access
api_name:
- Access.BoundObjectFrame.DisplayType
ms.assetid: 95213bcb-9751-b43c-9722-6326d0fa8f25
ms.date: 06/08/2017
---


# BoundObjectFrame.DisplayType Property (Access)

You can use the  **DisplayType** property to specify whether Microsoft Access displays an OLE object's content or an icon. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayType**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

For example, if the OLE object is a Microsoft Word document and you set this property to Content, the control displays the Word document; if you set this property to Icon, the control displays the Microsoft Word icon.

The  **DisplayType** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Content|**acOLEDisplayContent**|(Default) When the control contains an OLE object, the control displays the object's data, such as a document or spreadsheet.|
|Icon|**acOLEDisplayIcon**|When the control contains an OLE object, the control displays the object's icon.|
For a bound object frame, the  **DisplayType** property can be set either in Design view or in Form view or Datasheet view for new records while they are being added; it can be read in all views. For an unbound object frame or chart, the property can be set in the **Insert Object** dialog box when the object is created (the default setting is Content or, if you select the **Display As Icon** check box, the setting is Icon).

The  **DisplayType** property determines the default setting of the **Display As Icon** check box in the **Paste Special** dialog box, available by clicking **Paste Special** on the **Edit** menu, and the **Insert Object** dialog box, displayed when inserting an unbound object frame. When you display these dialog boxes in Form view, Datasheet view, or Design view, the **Display As Icon** check box is automatically selected if the **DisplayType** property is set to Icon. For example, you will see these boxes selected when using Visual Basic to set the control's **Action** property to **acOLEInsertObjDlg** or **acOLEPasteSpecialDlg**.

The  **DisplayType** property setting has no effect on the state of the **Display As Icon** check box in the **Object** dialog box when you insert an object into an unbound object frame. When you paste an object from the Clipboard, the **Display As Icon** check box reflects the state of the object on the Clipboard.

Changing the  **DisplayType** property of a bound object frame doesn't affect the display of existing objects in the control. However, it will affect new objects that you add to the control by using the **Object** command on the **Insert** menu.


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

