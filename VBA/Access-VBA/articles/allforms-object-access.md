---
title: AllForms Object (Access)
keywords: vbaac10.chm12683
f1_keywords:
- vbaac10.chm12683
ms.prod: access
api_name:
- Access.AllForms
ms.assetid: b90616b9-90fc-bb51-6bfa-b149dece0f1b
ms.date: 06/08/2017
---


# AllForms Object (Access)

The  **AllForms** collection contains an **[AccessObject](accessobject-object-access.md)** object for each form in the **[CurrentProject](currentproject-object-access.md)** or **[CodeProject](http://msdn.microsoft.com/library/70b71f57-df23-2cf7-23f5-147053a8ec26%28Office.15%29.aspx)** object.


## Remarks

The  **CurrentProject** and **CodeProject** object has an **AllForms** collection containing **AccessObject** objects that describe instances of all the forms in the database. For example, you can enumerate the **AllForms** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllForms** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllForms** collection, it's better to refer to the form by name because a form's collection index may change.

The  **AllForms** collection is indexed beginning with zero. If you refer to a form by its index, the first form is AllForms(0), the second form is AllForms(1), and so on.


 **Note**  To list all open forms in the database, use the  **[IsLoaded](http://msdn.microsoft.com/library/5e68398c-8a95-f3e1-87ec-e2d637f34429%28Office.15%29.aspx)** property of each **AccessObject** object in the **AllForms** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a form.

You can't add or delete an  **AccessObject** object from the **AllForms** collection.


## Example

The following example prints the name of each open  **AccessObject** object in the **AllForms** collection.


```
Sub AllForms() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentProject 
    ' Search for open AccessObject objects in AllForms collection. 
    For Each obj In dbs.AllForms 
        If obj.IsLoaded = True Then 
            ' Print name of obj. 
            Debug.Print obj.Name 
        End If 
    Next obj 
End Sub
```

The following example shows how to prevent a user form opening a particular form directly from the Navigation Pane.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)


```
'Don't let this form be opened from the Navigator
If Not CurrentProject.AllForms(cFormUsage).IsLoaded Then
    MsgBox "This form cannot be opened from the Navigation Pane.", _
        vbInformation + vbOKOnly, "Invalid form usage"
    Cancel = True
    Exit Sub
End If
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/fc74b94a-8a5a-a2b9-e534-5530d11d2907%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/1540145e-541d-10fc-249b-9fadc6861a11%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/07536c98-57e1-8660-b35e-0e79e4e797cb%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/fa16ed80-9eb2-7bd8-fdc6-a8c9a8eb7ea0%28Office.15%29.aspx)|

## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[AllForms Object Members](http://msdn.microsoft.com/library/a508646e-4478-fdfb-b1b5-177af651b73f%28Office.15%29.aspx)
