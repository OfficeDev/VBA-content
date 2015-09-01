
# Application.Explorers Property (Outlook)

 **Last modified:** July 28, 2015

Returns an  ** [Explorers](8398532a-1fad-7390-6778-109ac5e6c67c.md)**collection object that contains the  ** [Explorer](026591e5-049f-503a-4166-34e6dbc225fb.md)**objects representing all open explorers. Read-only.

## Syntax

 _expression_. **Explorers**

 _expression_A variable that represents an  **Application** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the number of explorer windows that are open.


```
Private Sub CountExplorers() 
 
 MsgBox "There are " &amp; _ 
 
 Application.Explorers.Count &amp; " Explorers." 
 
End Sub
```

The following VBA example uses the  ** [Count](ea7a19d2-6261-ce07-97f3-ebe95489a265.md)**property and  ** [Item](981b107a-14d7-2dd3-6449-2737b2801c3c.md)**method of the  ** [Selection](0b06a3ce-0445-db8f-e6e8-bb7bd469c50f.md)**collection returned by the  **Selection** property to display the senders of all mail items selected in the explorer that displays the **Inbox**. To run this example, you need to have at least one mail item selected in the explorer displaying the Inbox. You might receive an error if you select items other than a mail item such as task request as the  **SenderName** property does not exist for a ** [TaskRequestItem](2908a28a-634c-e786-aa53-f3e32038b727.md)** object.




```
Sub GetSelectedItems() 
 
 Dim myOlExp As Outlook.Explorer 
 
 Dim myOlSel As Outlook.Selection 
 
 Dim MsgTxt As String 
 
 Dim x As Integer 
 
 
 
 MsgTxt = "You have selected items from: " 
 
 Set myOlExp = Application.Explorers.Item(1) 
 
 If myOlExp = "Inbox" Then 
 
 Set myOlSel = myOlExp.Selection 
 
 For x = 1 To myOlSel.Count 
 
 MsgTxt = MsgTxt &amp; myOlSel.Item(x).SenderName &amp; ";" 
 
 Next x 
 
 MsgBox MsgTxt 
 
End If 
 
End Sub
```


## See also


#### Concepts


 [Application Object](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)
#### Other resources


 [Application Object Members](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)
