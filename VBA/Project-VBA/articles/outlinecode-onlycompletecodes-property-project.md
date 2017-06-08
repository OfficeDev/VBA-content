---
title: OutlineCode.OnlyCompleteCodes Property (Project)
ms.prod: project-server
api_name:
- Project.OutlineCode.OnlyCompleteCodes
ms.assetid: eb0b8dc2-2cb8-a86b-2711-fa4c6f215971
ms.date: 06/08/2017
---


# OutlineCode.OnlyCompleteCodes Property (Project)

 **True** if only outline codes with values at all levels of the code mask can be used. Read/write **Boolean**.


## Syntax

 _expression_. **OnlyCompleteCodes**

 _expression_ A variable that represents an **OutlineCode** object.


## Remarks

For enterprise text fields with a lookup table,  **OnlyCompleteCodes** is always **False** and non-writeable.


## Example

The following example adds a custom outline code to store the location of resources and configures the outline code such that only the full name of a code can be associated with a resource.


 **Note**  The  **OnlyCompleteCodes** property can be set only after the lookup table contains entries. If you try to set **OnlyCompleteCodes** before creating lookup table entries, the result is run-time error 7, "Out of memory."


```vb
Sub CreateLocationOutlineCode() 
 
 Dim objOutlineCode As OutlineCode 
 
 Set objOutlineCode = ActiveProject.OutlineCodes.Add( _ 
 pjCustomResourceOutlineCode1, "Location") 
 
 DefineLocationCodeMask objOutlineCode.CodeMask 
 EditLocationLookupTable objOutlineCode.LookupTable 
 
 objOutlineCode.OnlyCompleteCodes = True 
 
End Sub 
 
 
Sub DefineLocationCodeMask(objCodeMask As CodeMask) 
 objCodeMask.Add _ 
 Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
 Length:=2, Separator:="." 
 
 objCodeMask.Add _ 
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


