---
title: NewFile Object (Office)
keywords: vbaof11.chm235000
f1_keywords:
- vbaof11.chm235000
ms.prod: office
api_name:
- Office.NewFile
ms.assetid: 6f53ced5-4488-b67f-ca1f-729aeb790eb1
ms.date: 06/08/2017
---


# NewFile Object (Office)

Represents items listed on the  **New** _Item_ task pane available in several Microsoft Office applications.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

The following table shows the property to use to access the  **NewFile** object in each of the applications.


## Example

Use the  **Add** method to add a new item to the **New** _Item_ task pane. The following example adds an item to Word's **New Document** task pane.


```
Sub AddNewDocToTaskPane() 
    Application.NewDocument.Add FileName:="C:\NewDocument.doc", _ 
        Section:=msoNew, DisplayName:="New Document" 
    CommandBars("Task Pane").Visible = True  
End Sub
```

Use the  **Remove** method to remove an item from the **New** _Item_ task pane. The following example removes the document added in the above example from Word's **New Document** task pane.




```
Sub RemoveDocFromTaskPane() 
    Application.NewDocument.Remove FileName:="C:\NewDocument.doc", _ 
        Section:=msoNew, DisplayName:="New Document" 
    CommandBars("Task Pane").Visible = True  
End Sub
```


 **Note**  


 **Note**  The examples below are for Word, but you can change the  **NewDocument** property for any of the properties listed above and use the code in the corresponding application.


## Methods



|**Name**|
|:-----|
|[Add](newfile-add-method-office.md)|
|[Remove](newfile-remove-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](newfile-application-property-office.md)|
|[Creator](newfile-creator-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
