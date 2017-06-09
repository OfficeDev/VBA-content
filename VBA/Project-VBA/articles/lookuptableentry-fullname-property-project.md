---
title: LookupTableEntry.FullName Property (Project)
keywords: vbapj.chm132391
f1_keywords:
- vbapj.chm132391
ms.prod: project-server
api_name:
- Project.LookupTableEntry.FullName
ms.assetid: e1181061-5d49-7ae9-360f-1c397d744422
ms.date: 06/08/2017
---


# LookupTableEntry.FullName Property (Project)

Gets the full name for the specified level and parent levels of the  **LookupTableEntry** for the outline code, complete with the separator string between the levels. Read-only **String**.


## Syntax

 _expression_. **FullName**

 _expression_ A variable that represents a **LookupTableEntry** object.


## Example

The  **CreateLocationOutlineCode** macro example sets three **LookupTableEntry** levels for a custom task outline code named **Location**. After the **CreateLocationOutlineCode** macro is executed, entering the following line in the **Immediate** window of the Visual Basic Editor (VBE) returns the result shown.


```
Print ActiveProject.OutlineCodes.Item(1).LookupTable.Item(4).FullName 
WA.KING.RED
```

Following is the  **CreateLocationOutlineCode** macro.




```vb
Sub CreateLocationOutlineCode() 
    Dim objOutlineCode As OutlineCode 
    On Error GoTo ErrorHandler 
 
    Set objOutlineCode = ActiveProject.OutlineCodes.Add( _
        pjCustomTaskOutlineCode1, "Location") 
 
    objOutlineCode.OnlyLookUpTableCodes = True 
 
    DefineLocationCodeMask objOutlineCode.CodeMask 
    EditLocationLookupTable objOutlineCode.LookupTable 
 End 
 
ErrorHandler: 
    MsgBox "CreateLocationOutlineCode(): Error Number: " 
 &; Err.Number &; _ 
    vbCrLf &; " Error Description: " &; Err.Description 
End Sub 
 
 
Sub DefineLocationCodeMask(objCodeMask As CodeMask) 
    objCodeMask.Add _ 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _
        Length:=2, Separator:="." 
 
    objCodeMask.Add 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _
        Separator:="." 
 
    objCodeMask.Add _
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _
        Length:=3, Separator:="." 
End Sub 
 
 
Sub EditLocationLookupTable(objLookupTable As LookupTable) 
    Dim objStateEntry As LookupTableEntry 
    Dim objCountyEntry As LookupTableEntry 
    Dim objCityEntry As LookupTableEntry 
 
    Set objStateEntry = objLookupTable.AddChild("WA") 
    objStateEntry.Description = "Washington" 
 
    Set objCountyEntry = objLookupTable.AddChild("KING", _
        objStateEntry.UniqueID) 
    objCountyEntry.Description = "King County" 
 
    Set objCityEntry = objLookupTable.AddChild("SEA", _
        objCountyEntry.UniqueID) 
    objCityEntry.Description = "Seattle" 
 
    Set objCityEntry = objLookupTable.AddChild("RED", _
        objCountyEntry.UniqueID) 
    objCityEntry.Description = "Redmond" 
 
    Set objCityEntry = objLookupTable.AddChild("KIR", _
        objCountyEntry.UniqueID) 
    objCityEntry.Description = "Kirkland" 
End Sub
```


