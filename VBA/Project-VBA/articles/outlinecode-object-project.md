---
title: OutlineCode Object (Project)
ms.prod: project-server
api_name:
- Project.OutlineCode
ms.assetid: 8f75bdd3-ed5b-ed0f-9c3c-85af3a21580c
ms.date: 06/08/2017
---


# OutlineCode Object (Project)


 

Represents a local outline code in Project. The  **OutlineCode** object is a member of the **[OutlineCodes](outlinecodes-object-project.md)** collection.
 
 **Using the OutlineCode Object**
 
The following example adds a custom outline code to store the location of resources and configures the outline code so that only values specified in the lookup table can be associated with a resource. 
 



```
Sub CreateLocationOutlineCode() 
    Dim objOutlineCode As OutlineCode 
 
    Set objOutlineCode = ActiveProject.OutlineCodes.Add( _
        pjCustomResourceOutlineCode1, "Location") 
 
    objOutlineCode.OnlyLookUpTableCodes = True 
 
    DefineLocationCodeMask objOutlineCode.CodeMask 
    EditLocationLookupTable objOutlineCode.LookupTable 
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


## Remarks

An outline code is a type of local custom field that has a hierarchical text lookup table. Enterprise custom fields of type  **Text** that have hierarchical lookup tables act as outline codes. Use the **[OutlineCodes](project-outlinecodes-property-project.md)** property to return an **OutlineCodes** collection. Use the **[Add](outlinecodes-add-method-project.md)** method to add a local outline code to the **OutlineCodes** collection. To add an enterprise custom field, you must use Project Web App or the Project Server Interface (PSI).
 

 

## Methods



|**Name**|
|:-----|
|[Delete](outlinecode-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](outlinecode-application-property-project.md)|
|[CodeMask](outlinecode-codemask-property-project.md)|
|[DefaultValue](outlinecode-defaultvalue-property-project.md)|
|[FieldID](outlinecode-fieldid-property-project.md)|
|[Index](outlinecode-index-property-project.md)|
|[LinkedFieldID](outlinecode-linkedfieldid-property-project.md)|
|[LookupTable](outlinecode-lookuptable-property-project.md)|
|[MatchGeneric](outlinecode-matchgeneric-property-project.md)|
|[Name](outlinecode-name-property-project.md)|
|[OnlyCompleteCodes](outlinecode-onlycompletecodes-property-project.md)|
|[OnlyLeaves](outlinecode-onlyleaves-property-project.md)|
|[OnlyLookUpTableCodes](outlinecode-onlylookuptablecodes-property-project.md)|
|[Parent](outlinecode-parent-property-project.md)|
|[RequiredCode](outlinecode-requiredcode-property-project.md)|
|[SortOrder](outlinecode-sortorder-property-project.md)|

