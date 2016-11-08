
# WizardProperties Object (Publisher)

Represents the settings available in a publication design or in a Design Gallery object's wizard.
 


## Example

Use the  **[Properties](9f9811b3-10ee-d429-c5a2-8223349525f2.md)** property with a **Wizard** object to return a **WizardProperties** collection. The following example reports on the publication design associated with the active publication, displaying its name and current settings.
 

 

```
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 MsgBox "Publication Design associated with " _ 
 &amp; "current publication: " .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Note**  Depending on the language version of Microsoft Publisher that you are using, you may receive an error when using the above code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see  **[Wizard Object](c0a64ee9-d1fa-6dc7-5221-ff2d32874ea0.md)**.
 


## Methods



|**Name**|
|:-----|
|[FindPropertyById](9d13ffa2-f251-0e7d-2f36-c747413143d0.md)|

## Properties



|**Name**|
|:-----|
|[Application](2532d645-f317-90cb-6d78-d631bc116582.md)|
|[Count](835f3467-ec89-54d2-c685-3021e6267121.md)|
|[Item](e3f6732f-d093-4ccd-7c20-9fc357c0a8f5.md)|
|[Parent](b6e87015-67c9-834c-fe38-1dddee08e40a.md)|
