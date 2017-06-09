---
title: ObjectFrame.Action Property (Access)
keywords: vbaac10.chm11610
f1_keywords:
- vbaac10.chm11610
ms.prod: access
api_name:
- Access.ObjectFrame.Action
ms.assetid: 042d3418-fe67-c4cc-60b1-dc3b373b8d4f
ms.date: 06/08/2017
---


# ObjectFrame.Action Property (Access)

You can use the  **Action** property in Visual Basic to specify the operation to perform on an OLE object. Read/write **Integer**.


## Syntax

 _expression_. **Action**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

The  **Action** property uses the following settings.



|**Constant**|**Description**|
|:-----|:-----|
|**acOLECreateEmbed** (0)|Creates an embedded object. To use this setting, you must first set the control's **OLETypeAllowed** property to **acOLEEmbedded** or **acOLEEither**. Set the **Class** property to the type of OLE object you want to create. You can use the **SourceDoc** property to use an existing file as a template.|
|**acOLECreateLink** (1)|Creates a linked OLE object from the contents of a file. To use this setting, you must first set the control's  **OLETypeAllowed** and **SourceDoc** properties. Set the **OLETypeAllowed** property to **acOLELinked** or **acOLEEither**. The **SourceDoc** property specifies the file used to create the OLE object. You can also set the control's **SourceItem** property (for example, to specify a row-and-column range if the object you're creating is a Microsoft Excel worksheet). When you create an OLE object by using this setting, the control displays a metafile graphic image of the file specified by the control's **SourceDoc** property. If you save the OLE object, only the link information, such as the name of the application that supplied the object and the name of the linked file, is saved because the control contains an image of the data but no source data.|
|**acOLECopy** (4)|Copies the object to the Clipboard. When you copy an OLE object to the Clipboard, all the data and link information associated with the object is placed on the Clipboard as well. You can copy both linked and embedded objects onto the Clipboard. |
|**acOLEPaste** (5)|Pastes data from the Clipboard to the control. If the paste operation is successful, the control's  **OLEType** property is set to **acOLELinked** or **acOLEEmbedded**. If the paste operation isn't successful, the **OLEType** property is set to **acOLENone**.|
|**acOLEUpdate** (6)|Retrieves the current data from the application that supplied the object and displays that data as a metafile graphic in the control.|
|**acOLEActivate** (7)|Opens an OLE object for an operation, such as editing. To use this setting, you must first set the control's  **Verb** property. The **Verb** property specifies the operation to perform when the OLE object is activated.|
|**acOLEClose** (9)|Closes an OLE object and ends the connection with the application that supplied the object. This setting applies to embedded objects only. Using this setting is equivalent to clicking  **Close** on the object's **Control** menu.|
|**acOLEDelete** (10)|Deletes the specified OLE object and frees the associated memory. This setting enables you to explicitly delete an OLE object. Objects are automatically deleted when a form is closed or when the object is updated to a new object. You can't use the  **Action** property to delete a bound OLE object from its underlying table or query.|
|**acOLEInsertObjDlg** (14)|Displays the  **Insert Object** dialog box. In Form view or Datasheet view, you display this dialog box to enable the user to create a new object or to link or embed an existing object. You can use the control's **OLETypeAllowed** property to determine the type of object the user can create (with the constant **acOLELinked**, **acOLEEmbedded**, or **acOLEEither** ) by using this dialog box.|
|**acOLEPasteSpecialDlg** (15)|Displays the  **Paste Special** dialog box. In Form view or Datasheet view, you display this dialog box to enable the user to paste an object from the Clipboard. The dialog box provides several options, including pasting either a linked or embedded object. You can use the control's **OLETypeAllowed** property to determine the type of object that can be pasted (with the constant **acOLELinked**, **acOLEEmbedded**, or **acOLEEither** ) by using this dialog box.|
|**acOLEFetchVerbs** (17)|Updates the list of verbs an OLE object supports. To display the list of verbs, use the  **ObjectVerbs** and **ObjectVerbsCount** properties.|
The  **Action** property isn't available in Design view but can be read or set in other views.

When a control's  **Enabled** property is set to No or its **Locked** property is set to Yes, you can't use some **Action** property settings. The following table indicates which settings are allowed or not allowed under these conditions.



|**Setting**|**Enabled = No**|**Locked = Yes**|
|:-----|:-----|:-----|
|**acOLECreateEmbed** (0)|Not allowed|Not allowed|
|**acOLECreateLink** (1)|Not allowed|Not allowed|
|**acOLECopy** (4)|Allowed|Allowed|
|**acOLEPaste** (5)|Not allowed|Not allowed|
|**acOLEUpdate** (6)|Not allowed|Not allowed|
|**acOLEActivate** (7)|Allowed|Allowed|
|**acOLEClose** (9)|Not allowed|Allowed|
|**acOLEDelete** (10)|Not allowed|Not allowed|
|**acOLEInsertObjDlg** (14)|Not allowed|Not allowed|
|**acOLEPasteSpecialDlg** (15)|Not allowed|Not allowed|
|**acOLEFetchVerbs** (17)|Not allowed|Allowed|

## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

