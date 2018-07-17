---
title: OutlineCodes Object (Project)
ms.prod: project-server
ms.assetid: a2e6d0c7-0741-91c6-61aa-f4bcc299e66f
ms.date: 06/08/2017
---


# OutlineCodes Object (Project)

Contains a collection of  **[OutlineCode](outlinecode-object-project.md)** objects.
 


## Remarks

An outline code is a type of local custom field that has a hierarchical text lookup table. Enterprise custom fields of type  **Text** that have hierarchical lookup tables act as outline codes. Use the **[OutlineCodes](project-outlinecodes-property-project.md)** property to return an **OutlineCodes** collection. Use the **[Add](outlinecodes-add-method-project.md)** method to add a local outline code to the **OutlineCodes** collection. To add an enterprise custom field, you must use Project Web App or the Project Server Interface (PSI).
 

 

## Example

 **Using the OutlineCodes Collection Object**
 

 
The following example adds a custom outline code to store the location of resources and configures the outline code such that only values specified in the lookup table can be associated with a resource. 
 

 

 **Note**  The  **OnlyLookUpTableCodes** property can be set only after the lookup table contains entries. If you try to set **OnlyLookUpTableCodes** before creating lookup table entries, the result is run-time error 7, "Out of memory."
 




```
Sub CreateLocationOutlineCode() 

 

 Dim objOutlineCode As OutlineCode 

 

 Set objOutlineCode = ActiveProject.OutlineCodes.Add( _ 

 pjCustomResourceOutlineCode1, "Location") 

 

 DefineLocationCodeMask objOutlineCode.CodeMask 

 EditLocationLookupTable objOutlineCode.LookupTable 

 

 objOutlineCode.OnlyLookUpTableCodes = True 

 

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


## Methods



|**Name**|
|:-----|
|[Add](outlinecodes-add-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](outlinecodes-application-property-project.md)|
|[Count](outlinecodes-count-property-project.md)|
|[Item](outlinecodes-item-property-project.md)|
|[Parent](outlinecodes-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
