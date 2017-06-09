---
title: Change a Content Control from One Type to Another
ms.prod: word
ms.assetid: e89924d4-3088-2e9a-0797-4553e2ff5ff0
ms.date: 06/08/2017
---


# Change a Content Control from One Type to Another

The content controls that you can create in documents in Word 2007 and later are extremely flexible. In most cases, you can easily switch a content control from one type to another. For example, if you have a date content control, you can change it to be a text content control, or if you have a text content control, you can change it to be a rich-text content control. To do this, you use the  **[Type](contentcontrol-type-property-word.md)** property and set it to a **[WdContentControlType](wdcontentcontroltype-enumeration-word.md)** constant.


 **Note**  Depending on the contents of a content control, you may not be able to change the content control type. For example, if you have a rich-text content control that contains formatted text, you may not be able to change the control to be a plain-text content control. In this case, Word raises a run-time error. 


The objects used in this sample are:


-  [ContentControl](contentcontrol-object-word.md)
    
-  [ContentControls](contentcontrols-object-word.md)
    
-  [ContentControlListEntries](contentcontrollistentries-object-word.md)
    
The following example inserts a new date content control that contains the current date and then changes it to be a text content control.



```vb
Sub ChangeTypeOfControl() 
 Dim objCC As ContentControl 
 Dim strDate As Date 
 
 strDate = Date 
 Set objCC = Selection.ContentControls.Add(wdContentControlDate) 
 objCC.Range.Text = strDate 
 
 objCC.Type = wdContentControlText 
End Sub
```

The following example inserts a drop-down list content control and then changes it to be a rich-text content control.



```vb
Sub ChangeContentControlType() 
 Dim objCC As ContentControl 
 
 Set objCC = ActiveDocument.ContentControls.Add(Type:=wdContentControlDropdownList) 
 objCC.SetPlaceholderText Text:="My Favorite Animal" 
 
 'List entries 
 objCC.DropdownListEntries.Add "Cat" 
 objCC.DropdownListEntries.Add "Dog" 
 objCC.DropdownListEntries.Add "Horse" 
 objCC.DropdownListEntries.Add "Monkey" 
 objCC.DropdownListEntries.Add "Snake" 
 objCC.DropdownListEntries.Add "Other" 
 
 Stop 
 
 ' Switch to view the new content control in the active document. 
 ' Notice that the content control is a drop-down list. 
 
 objCC.Type = wdContentControlRichText 
 
 ' After running the above code, the content control is no longer 
 ' a drop-down; it is a text content control. Only the placeholder 
 ' text remains; Word removes the items in the list. 
End Sub
```


