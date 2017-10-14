---
title: Months.Count Property (Project)
ms.prod: project-server
api_name:
- Project.Months.Count
ms.assetid: c686777e-5540-5f1c-7e50-e5138b12e280
ms.date: 06/08/2017
---


# Months.Count Property (Project)

Gets the number of items in the  **Months** collection for a specified year from 1984 - 2149. Read-only **Integer**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Months** object.


## Examples

The following example in the  **Immediate** window of the VBE returns 12, the number of months in 2012. If you set the year to 1983 or 2150, the result is "Run-time error '1101'; the argument is not valid."


```vb
? activeproject.Resources(1).Calendar.Years(2012).Months.Count
```

The following example shows the use of the  **Count** property for the **Assignments** object. It prompts the user for the name of a resource and then assigns that resource to tasks without any resources.




```vb
Sub AssignResource()  
    Dim T As Task ' Task object used in For Each loop  
    Dim R As Resource ' Resource object used in For Each loop  
    Dim Rname As String ' Resource name  
    Dim RID As Long ' Resource ID  
  
    RID = 0  
    RName = InputBox$("Enter the name of a resource: ")  
  
    For Each R in ActiveProject.Resources  
        If R.Name = RName Then  
            RID = R.ID  
            Exit For  
        End If  
    Next R  
  
    If RID <> 0 Then  
        ' Assign the resource to tasks without any resources.  
        For Each T In ActiveProject.Tasks  
            If T.Assignments.Count = 0 Then  
                T.Assignments.Add ResourceID:=RID  
            End If
        Next T
    Else  
        MsgBox Prompt:=RName &; " is not a resource in this project.", buttons:=vbExclamation
    End If
End Sub
```


## See also


#### Concepts


[Months Collection Object](months-object-project.md)
