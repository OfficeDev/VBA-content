---
title: ObjectFrame.Object Property (Access)
keywords: vbaac10.chm11550
f1_keywords:
- vbaac10.chm11550
ms.prod: access
api_name:
- Access.ObjectFrame.Object
ms.assetid: db4354f5-a92d-7341-7823-a1c5f26d74b8
ms.date: 06/08/2017
---


# ObjectFrame.Object Property (Access)

You can use the  **Object** property in Visual Basic to return a reference to the ActiveX object that is associated with a linked or embedded OLE object in a control. By using this reference, you can access the properties or invoke the methods of the OLE object. Read-only **Object**.


## Syntax

 _expression_. **Object**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

The  **Object** property returns a reference to an ActiveX object. You can use the **Set** statement to assign this ActiveX object to an object variable. The type of object reference returned depends on which application created the OLE object.

When you embed or link an OLE object in a Microsoft Access form, you can set properties that determine the type of object and the behavior of the container control. However, you can't directly set or read the OLE object's properties or apply its methods, as you can when performing Automation. The  **Object** property returns a reference to an Automation object that represents the linked or embedded OLE object. By using this reference, you can change the OLE object by setting or reading its properties or applying its methods. For example, Microsoft Excel is an COM component that supports Automation. If you've embedded a Microsoft Excel worksheet in a Microsoft Access form, you can use the **Object** property to set a reference to the **Worksheet** object associated with that worksheet. You can then use any of the properties and methods of the **Worksheet** object.

For information on which properties and methods an ActiveX object supports, see the documentation for the application that was used to create the OLE object.


## Example

The following example uses the  **Object** property of an unbound object frame named OLE1. Customer name and address information is inserted in an embedded Microsoft Word document formatted as a form letter with placeholders for the name and address information and boilerplate text in the body of the letter. The procedure replaces the placeholder information for each record and prints the form letter. It doesn't save copies of the printed form letter.


```vb
Sub PrintFormLetter_Click() 
 Dim objWord As Object 
 Dim strCustomer As String, strAddress As String 
 Dim strCity As String, strRegion As String 
 
 ' Assign object property of control to variable. 
 Set objWord = Me!OLE1.Object.Application.Wordbasic 
 ' Assign customer address to variables. 
 strCustomer = Me!CompanyName 
 strAddress = Me!Address 
 strCity = Me!City &; ", " 
 If Not IsNull(Me!Region) Then 
 strRegion = Me!Region 
 Else 
 strRegion = Me!Country 
 End If 
 ' Activate ActiveX control. 
 Me!OLE1.Action = acOLEActivate 
 With objWord 
 .StartOfDocument 
 ' Go to first placeholder. 
 .LineDown 2 
 ' Highlight placeholder text. 
 .EndOfLine 1 
 ' Insert customer name. 
 .Insert strCustomer 
 ' Go to next placeholder. 
 .LineDown 
 .StartOfLine 
 ' Highlight placeholder text. 
 .EndOfLine 1 
 ' Insert address. 
 .Insert strAddress 
 ' Go to last placeholder. 
 .LineDown 
 .StartOfLine 
 ' Highlight placeholder text. 
 .EndOfLine 1 
 ' Insert City and Region. 
 .Insert strCity &; strRegion 
 .FilePrint 
 .FileClose 
 End With 
 Set objWord = Nothing 
End Sub
```


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

