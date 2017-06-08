---
title: Shape.Wizard Property (Publisher)
keywords: vbapb10.chm2228345
f1_keywords:
- vbapb10.chm2228345
ms.prod: publisher
api_name:
- Publisher.Shape.Wizard
ms.assetid: 89014daf-66dc-7913-0b0e-ac80f6e85791
ms.date: 06/08/2017
---


# Shape.Wizard Property (Publisher)

Returns a  **[Wizard](wizard-object-publisher.md)** object representing the publication design associated with the specified publication or the wizard associated with the specified Design Gallery object.


## Syntax

 _expression_. **Wizard**

 _expression_A variable that represents a  **Shape** object.


## Remarks

When accessing the  **Wizard** property from the **Document** or **Page** object, if the specified publication is not associated with any publication design, an error occurs. When accessing the **Wizard** property from the **Shape** or **ShapeRange** object, if the specified object is not a Design Gallery object, an error occurs.


## Example

The following example reports on the publication design associated with the active publication, displaying its name and current settings.


```vb
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 Debug.Print "Publication design associated with " _ 
 &; "current publication: " _ 
 &; .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Setting: " _ 
 &; .Name &; " = " &; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Note**  Depending on the language version of Publisher that you are using, you may receive an error when using the above code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see  **[Wizard](wizard-object-publisher.md)** object .


