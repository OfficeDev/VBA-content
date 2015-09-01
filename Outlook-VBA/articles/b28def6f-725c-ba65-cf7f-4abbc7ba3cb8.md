
# SelectNamesDialog.CcLabel Property (Outlook)

 **Last modified:** July 28, 2015

Returns or sets a  **String** for the text that appears on the **Cc** command button on the **Select Names** dialog box. Read/write.

## Syntax

 _expression_. **CcLabel**

 _expression_A variable that represents a  **SelectNamesDialog** object.


## Remarks

To provide an accelerator key for the recipient edit boxes, include an ampersand (&amp;) character in the label argument string, immediately before the character that serves as the access key. For example, if  **CcLabel** is the string "Local &amp;Attendees", users can press **ALT+A** to move the focus to the first recipient edit box.

If  **CcLabel** is not set, then the default value will be the localized string for "Cc". If the **CcLabel** is set to an empty string, then the **Cc** command button shows **-&gt;**. If the  **CcLabel** property contains more than 32 characters (including the ampersand (&amp;) access key), only the first 32 characters will be displayed in the command button.


## See also


#### Concepts


 [SelectNamesDialog Object](1522736a-3cad-9f1c-4da9-b52a3a01731c.md)
#### Other resources


 [SelectNamesDialog Object Members](0f5546af-f89a-8a8b-ced9-a2d646bf9634.md)
