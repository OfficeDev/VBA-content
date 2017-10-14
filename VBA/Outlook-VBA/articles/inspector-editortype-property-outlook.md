---
title: Inspector.EditorType Property (Outlook)
keywords: vbaol11.chm2963
f1_keywords:
- vbaol11.chm2963
ms.prod: outlook
api_name:
- Outlook.Inspector.EditorType
ms.assetid: b19e552b-1e8a-8915-f793-396860910f40
ms.date: 06/08/2017
---


# Inspector.EditorType Property (Outlook)

Returns an  **[OlEditorType](oleditortype-enumeration-outlook.md)** constant indicating the type of editor. Read-only.


## Syntax

 _expression_ . **EditorType**

 _expression_ A variable that represents an **Inspector** object.


## Remarks

Since Microsoft Office Outlook 2007, the  **EditorType** property always returns **olEditorWord** .


## Example

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  **[Open](mailitem-open-event-outlook.md)** event to access the **[HTMLBody](mailitem-htmlbody-property-outlook.md)** property of an item. This sets the **[EditorType](inspector-editortype-property-outlook.md)** property of the item's **[Inspector](inspector-object-outlook.md)** to **olEditorHTML** . If this code is placed in the Script Editor of a form in design mode, the message boxes during run time will reflect the change in the **EditorType** as the body of the form changes. The final message box utilizes the **[ScriptText](formdescription-scripttext-property-outlook.md)** property to display all the VBScript code in the Script Editor.


```vb
Function Item_Open() 
 'Set the HTMLBody of the item. 
 Item.HTMLBody = "<HTML><H2>My HTML page.</H2><BODY>My body.</BODY></HTML>" 
 'Item displays HTML message. 
 Item.Display 
 'MsgBox shows EditorType is 2 which represents the HTML editor type 
 MsgBox "HTMLBody EditorType is " &; Item.GetInspector.EditorType 
 'Access the Body and show 
 'the text of the Body. 
 MsgBox "This is the Body: " &; Item.Body 
 'After accessing, EditorType 
 'is still 2. 
 MsgBox "After accessing, the EditorType is " &; Item.GetInspector.EditorType 
 'Set the item's Body property. 
 Item.Body = "Back to default body." 
 'After setting the Body, EditorType is 
 'still the same. 
 MsgBox "After setting, the EditorType is " &; Item.GetInspector.EditorType 
End Function
```


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

