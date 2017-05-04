---
title: NewFile Object (Office)
keywords: vbaof11.chm235000
f1_keywords:
- vbaof11.chm235000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.NewFile
ms.assetid: 6f53ced5-4488-b67f-ca1f-729aeb790eb1
---


# NewFile Object (Office)

Represents items listed on the  **New** _Item_ task pane available in several Microsoft Office applications.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

The following table shows the property to use to access the  **NewFile** object in each of the applications.


## Example

Use the  **Add** method to add a new item to the **New** _Item_ task pane. The following example adds an item to Word's **New Document** task pane.


```vb
Sub AddNewDocToTaskPane() 
    Application.NewDocument.Add FileName:="C:\NewDocument.doc", _ 
        Section:=msoNew, DisplayName:="New Document" 
    CommandBars("Task Pane").Visible = True  
End Sub
```

Use the  **Remove** method to remove an item from the **New** _Item_ task pane. The following example removes the document added in the above example from Word's **New Document** task pane.




```vb
Sub RemoveDocFromTaskPane() 
    Application.NewDocument.Remove FileName:="C:\NewDocument.doc", _ 
        Section:=msoNew, DisplayName:="New Document" 
    CommandBars("Task Pane").Visible = True  
End Sub
```


 **Note**  


 **Note**  The examples below are for Word, but you can change the  **NewDocument** property for any of the properties listed above and use the code in the corresponding application.


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

