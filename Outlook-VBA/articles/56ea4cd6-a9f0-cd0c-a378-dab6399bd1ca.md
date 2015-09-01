
# FormDescription.ScriptText Property (Outlook)

 **Last modified:** July 28, 2015

Returns a  **String** containing all the VBScript code in the form's Script Editor. Read-only.

## Syntax

 _expression_. **ScriptText**

 _expression_A variable that represents a  ** [FormDescription](c88f92c4-4cac-84b3-6118-1150d42d7cff.md)** object.


## Example

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  ** [Open](656c16f7-d561-a8f7-e859-9ac24f357769.md)**event to access the  ** [HTMLBody](c340fe05-9a99-3a32-3d6b-f2f7a568b299.md)**property of a  ** [MailItem](14197346-05d2-0250-fa4c-4a6b07daf25f.md)**. This sets the  ** [EditorType](b19e552b-1e8a-8915-f793-396860910f40.md)**property of the  **MailItem**'s  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)**to  **olEditorHTML**. When the  **MailItem**'s  ** [Body](578567b1-893b-db4e-dddb-f3c237952c03.md)**property is set, the  **EditorType** property is changed to the default. For example, if the default e-mail editor is set to RTF, the **EditorType** is set to **olEditorRTF**. If this code is placed in the Script Editor of a form in design mode, the message boxes during run time will reflect the change in the  **EditorType** as the body of the form changes. The final message box uses the **Script Text**property to display all the VBScript code in the Script Editor.


```
Function Item_Open() 
 
 'Set the HTMLBody of the item. 
 
 Item.HTMLBody = "<HTML><H2>My HTML page.</H2><BODY>My body.</BODY></HTML>" 
 
 'Item displays HTML message. 
 
 Item.Display 
 
 'MsgBox shows EditorType is 2. 
 
 MsgBox "HTMLBody EditorType is " &amp; Item.GetInspector.EditorType 
 
 'Access the Body and show 
 
 'the text of the Body. 
 
 MsgBox "This is the Body: " &amp; Item.Body 
 
 'After accessing, EditorType 
 
 'is still 2. 
 
 MsgBox "After accessing, the EditorType is " &amp; Item.GetInspector.EditorType 
 
 'Set the item's Body property. 
 
 Item.Body = "Back to default body." 
 
 'After setting, EditorType is 
 
 'now back to the default. 
 
 MsgBox "After setting, the EditorType is " &amp; Item.GetInspector.EditorType 
 
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


 [FormDescription Object](c88f92c4-4cac-84b3-6118-1150d42d7cff.md)
#### Other resources


 [FormDescription Object Members](664724e9-e74b-32ad-93e4-8d4cb27b3082.md)
