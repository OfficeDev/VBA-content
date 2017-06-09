---
title: FormDescription.ScriptText Property (Outlook)
keywords: vbaol11.chm197
f1_keywords:
- vbaol11.chm197
ms.prod: outlook
api_name:
- Outlook.FormDescription.ScriptText
ms.assetid: 56ea4cd6-a9f0-cd0c-a378-dab6399bd1ca
ms.date: 06/08/2017
---


# FormDescription.ScriptText Property (Outlook)

Returns a  **String** containing all the VBScript code in the form's Script Editor. Read-only.


## Syntax

 _expression_ . **ScriptText**

 _expression_ A variable that represents a **[FormDescription](formdescription-object-outlook.md)** object.


## Example

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  **[Open](mailitem-open-event-outlook.md)** event to access the **[HTMLBody](mailitem-htmlbody-property-outlook.md)** property of a **[MailItem](mailitem-object-outlook.md)** . This sets the **[EditorType](inspector-editortype-property-outlook.md)** property of the **MailItem** 's **[Inspector](inspector-object-outlook.md)** to **olEditorHTML** . When the **MailItem** 's **[Body](mailitem-body-property-outlook.md)** property is set, the **EditorType** property is changed to the default. For example, if the default e-mail editor is set to RTF, the **EditorType** is set to **olEditorRTF** . If this code is placed in the Script Editor of a form in design mode, the message boxes during run time will reflect the change in the **EditorType** as the body of the form changes. The final message box uses the **Script Text** property to display all the VBScript code in the Script Editor.


```vb
Function Item_Open() 
 
 'Set the HTMLBody of the item. 
 
 Item.HTMLBody = "<HTML><H2>My HTML page.</H2><BODY>My body.</BODY></HTML>" 
 
 'Item displays HTML message. 
 
 Item.Display 
 
 'MsgBox shows EditorType is 2. 
 
 MsgBox "HTMLBody EditorType is " &; Item.GetInspector.EditorType 
 
 'Access the Body and show 
 
 'the text of the Body. 
 
 MsgBox "This is the Body: " &; Item.Body 
 
 'After accessing, EditorType 
 
 'is still 2. 
 
 MsgBox "After accessing, the EditorType is " &; Item.GetInspector.EditorType 
 
 'Set the item's Body property. 
 
 Item.Body = "Back to default body." 
 
 'After setting, EditorType is 
 
 'now back to the default. 
 
 MsgBox "After setting, the EditorType is " &; Item.GetInspector.EditorType 
 
 'Access the items's 
 
 'FormDescription object. 
 
 Set myForm = Item.FormDescription 
 
 'Display all the code 
 
 'in the Script Editor. 
 
 MsgBox myForm.ScriptText 
 
End Function
```


## See also


#### Concepts


[FormDescription Object](formdescription-object-outlook.md)

