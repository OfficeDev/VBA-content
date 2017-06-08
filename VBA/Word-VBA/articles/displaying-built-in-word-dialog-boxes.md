---
title: Displaying Built-in Word Dialog Boxes
keywords: vbawd10.chm5210531
f1_keywords:
- vbawd10.chm5210531
ms.prod: word
ms.assetid: abe465f9-09a1-72ea-2e2d-9de14fc02434
ms.date: 06/08/2017
---


# Displaying Built-in Word Dialog Boxes

This topic contains the following information and examples:


-  [Showing a built-in dialog box](#item1)
    
-  [Returning and changing dialog box settings](#item2)
    
-  [Checking how a dialog box was closed](#item4)
    

## Showing a built-in dialog box

You can display a built-in dialog box to get user input or to control Word by using Visual Basic for Applications (VBA). The  **[Show](dialog-show-method-word.md)** method of the **[Dialog](dialog-object-word.md)** object displays and executes any action taken in a built-in Word dialog box. To access a particular built-in Word dialog box, you specify a **[WdWordDialog](wdworddialog-enumeration-word.md)** constant with the **[Dialogs](application-dialogs-property-word.md)** property. For example, the following macro instruction displays the **Open** dialog box ( **wdDialogFileOpen**).


```vb
Sub ShowOpenDialog() 
 Dialogs(wdDialogFileOpen).Show 
End Sub
```

If a file is selected and  **OK** is clicked, the file is opened (the action is executed). The following example displays the **Print** dialog box ( **wdDialogFilePrint**).




```vb
Sub ShowPrintDialog() 
 Dialogs(wdDialogFilePrint).Show 
End Sub
```

Set the  **[DefaultTab](dialog-defaulttab-property-word.md)** property to access a particular tab in a Word dialog box. The following example displays the **Page Border** tab in the **Borders and Shading** dialog box.




```vb
Sub ShowBorderDialog() 
 With Dialogs(wdDialogFormatBordersAndShading) 
 .DefaultTab = wdDialogFormatBordersAndShadingTabPageBorder 
 .Show 
 End With 
End Sub
```


 **Note**  You can also use the VBA properties in Word to display the user information without displaying the dialog box. The following example uses the  **[UserName](application-username-property-word.md)** property for the **[Application](application-object-word.md)** object to display the user name for the application without displaying the **User Information** dialog box.




```vb
Sub DisplayUserInfo() 
 MsgBox Application.UserName 
End Sub
```

If the user name is changed in the previous example, the change is not set in the dialog box. Use the  **[Execute](dialog-execute-method-word.md)** method to execute the settings in a dialog box without displaying the dialog box. The following example displays the **User Information** dialog box, and if the name is not an empty string, the settings are set in the dialog box by using the **Execute** method.




```vb
Sub ShowAndSetUserInfoDialogBox() 
 With Dialogs(wdDialogToolsOptionsUserInfo) 
 .Display 
 If .Name <> "" Then .Execute 
 End With 
End Sub
```


 **Note**  Use the VBA properties and methods in Word to set the user information without displaying the dialog box. The following code example changes the user name through the  **UserName** property of the **Application** object, and then it displays the **User Information** dialog box to show that the change has been made. Note that displaying the dialog box is not necessary to change the value of a dialog box.




```vb
Sub SetUserName() 
 Application.UserName = "Jeff Smith" 
 Dialogs(wdDialogToolsOptionsUserInfo).Display 
End Sub
```


## Returning and changing dialog box settings

It is not very efficient to use a  **Dialog** object to return or change a value for a dialog box when you can return or change it using a property or method. Also, in most, if not all, cases, when VBA code is used in place of accessing the **Dialog** object, code is simpler and shorter. Therefore, the following examples also include examples that use corresponding VBA properties to perform the same tasks.

Prior to returning or changing a dialog box setting using the  **[Dialog](dialog-object-word.md)** object, you need to identify the individual dialog box. This is done by using the **[Dialogs](dialogs-count-property-word.md)** property with a **[WdWordDialog](wdworddialog-enumeration-word.md)** constant. After you have instantiated a **Dialog** object, you can return or set options in the dialog box. The following example displays the right indent from the **Paragraphs** dialog box.




```vb
Sub ShowRightIndent() 
 Dim dlgParagraph As Dialog 
 Set dlgParagraph = Dialogs(wdDialogFormatParagraph) 
 MsgBox "Right indent = " &; dlgParagraph.RightIndent 
End Sub
```


 **Note**  You can use the VBA properties and methods of Word to display the right indent setting for the paragraph. The following example uses the  **[RightIndent](paragraphformat-rightindent-property-word.md)** property of the **[ParagraphFormat](paragraphformat-object-word.md)** object to display the right indent for the paragraph at the insertion point position.




```vb
Sub ShowRightIndexForSelectedParagraph() 
 MsgBox Selection.ParagraphFormat.RightIndent 
End Sub
```

Just as you can return dialog box settings, you can also set dialog box settings. The following example marks the  **Keep with next** check box in the **Paragraph** dialog box.




```vb
Sub SetKeepWithNext() 
 With Dialogs(wdDialogFormatParagraph) 
 .KeepWithNext = 1 
 .Execute 
 End With 
End Sub
```


 **Note**  You can also use the VBA properties and methods to change the right indent for the paragraph. The following example uses the  **[KeepWithNext](paragraphformat-keepwithnext-property-word.md)** property of the **ParagraphFormat** object to keep the selected paragraph with the following paragraph.




```vb
Sub SetKeepWithNextForSelectedParagraph() 
 Selection.ParagraphFormat.KeepWithNext = True 
End Sub
```


 **Note**  Use the  **[Update](dialog-update-method-word.md)** method to ensure that the dialog box values reflect the current values. It may be necessary to use the **Update** method if you define a dialog box variable early in your macro and later want to return or change the current settings.


## Checking how a dialog box was closed

The value returned by the  **Show** and **Display** methods indicates which button was clicked to close the dialog box. The following example displays the **Break** dialog box, and if **OK** is clicked, a message is displayed on the status bar.


```vb
Sub DialogBoxButtons() 
 If Dialogs(wdDialogInsertBreak).Show = -1 Then 
 StatusBar = "Break inserted" 
 End If 
End Sub
```

The following table describes the return values associated with buttons in dialogs boxes.



|**Return value**|**Description**|
|:-----|:-----|
|-2|The  **Close** button.|
|-1|The  **OK** button.|
|0 (zero)|The  **Cancel** button.|
|> 0 (zero)|A command button: 1 is the first button, 2 is the second button, and so on.|

