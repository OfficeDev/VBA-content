---
title: Change the Editability of a Content Control
ms.prod: word
ms.assetid: ec1856d6-c19a-cc73-1d9c-237935e69db1
ms.date: 06/08/2017
---


# Change the Editability of a Content Control

Sometimes you might want to limit how a user works with a content control. Word gives you several ways to do this. For example, you can limit the text that a user can insert into a content control, force a user to insert content into a content control, or lock a content control. There are two ways that you can lock a content control. One prohibits a user from deleting the content control. Another prohibits a user from editing the content control.

You can set these restrictions programmatically. For example, you might want to prohibit a user from editing a content control based on the value that a user inserts into another control. Use the  **[LockContentControl](contentcontrol-lockcontentcontrol-property-word.md)** property to prohibit a user from deleting a content control, and use the **[LockContents](contentcontrol-lockcontents-property-word.md)** property to prohibit a user from editing the contents of a content control.

The objects used in this sample are:


-  [ContentControl](contentcontrol-object-word.md)
    
-  [ContentControls](contentcontrols-object-word.md)
    
The following example uses the  **LockContentControl** property and the **LockContents** property to prohibit a user from deleting or editing the content control.



```vb
Sub LockcontentControl() 
 Dim objCC As ContentControl 
 
 Set objCC = ActiveDocument.ContentControls _ 
 .Add(wdContentControlRichText) 
 
 objCC.LockcontentControl = True 
 objCC.LockContents = True 
End Sub
```

Using these properties with one of the events for content controls, such as the  **ContentControlOnExit** event, gives you control over how content controls are used in documents and how users work with them.

