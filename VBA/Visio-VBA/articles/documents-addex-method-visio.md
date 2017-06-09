---
title: Documents.AddEx Method (Visio)
keywords: vis_sdr.chm10651450
f1_keywords:
- vis_sdr.chm10651450
ms.prod: visio
api_name:
- Visio.Documents.AddEx
ms.assetid: 4c287668-04b4-fb6c-2b1a-869d9d366991
ms.date: 06/08/2017
---


# Documents.AddEx Method (Visio)

Adds a new stencil or drawing to the  **Documents** collection, while permitting extra information to be passed in an argument.


## Syntax

 _expression_ . **AddEx**( **_FileName_** , **_MeasurementSystem_** , **_Flags_** , **_LangID_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The type or file name of the document to add; if you do not include a path, Microsoft Visio searches the folder or folders designated in the ** Application** object's **TemplatePaths** property and all published templates, including published third-party templates.|
| _MeasurementSystem_|Optional| **VisMeasurementSystem**|The measurement units to use in the new document. See Remarks for possible values.|
| _Flags_|Optional| **Long**|Flags that indicate how to open the new document. See Remarks for possible values.|
| _LangID_|Optional| **Long**|The language ID for the document. See Remarks.|

### Return Value

Document


## Remarks

The  **AddEx** method is similar to the **Add** method as it applies to the **Documents** collection, except that **AddEx** provides several additional arguments in which the caller can specify how the document is created.

The  _MeasurementSystem_ argument should be one of the following members of **VisMeasurementSystem** , which is declared in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visMSDefault**|0|Choose metric or US depending on regional options set in Control Panel.|
| **visMSMetric**|1|Metric measurement system.|
| **visMSUS**|2|US units measurement system.|
The  _Flags_ argument should be a combination of one or more of the following members of **VisOpenSaveArgs** , which is declared in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visAddDocked**|4|Adds a document in a docked window.|
| **visAddHidden**|64|Adds a document in a hidden window.|
| **visAddMacrosDisabled**|128|Adds a document with macros disabled.|
| **visAddMinimized**|16|Adds a document in a minimized window.|
| **visAddNoWorkspace**|256|Adds a document with no workspace information.|
| **visAddStencil**|512|Adds a new stencil file.|
The  _LangID_ argument should be one of the standard IDs used by Microsoft Windows to encode different language versions. For example, the language ID is &;H0409 for the U.S. version of Visio. To see a list of language IDs, search for "VERSIONINFO" in the Microsoft Platform SDK on MSDN.

To create a new drawing based on no template, pass a zero-length string ("") to the  **AddEx** method.

To create a new drawing based on a template, pass "templatename.vst" to the  **AddEx** method. Visio opens stencils that are part of the template's workspace and copies styles and other settings associated with the template to the new document. If the template file name is invalid, no document is returned and an error is generated.

To create a new stencil based on no stencil, pass ("vss").

To open a copy of a stencil, pass ("stencilname.vss").

To open a copy of a drawing, pass ("drawingname.vsd").




 **Note**  Opening a copy of a stencil or drawing is equivalent to selecting  **Open as Copy** in the **Open** list in the **Open** dialog box or using the **OpenEx** method with the **visOpenCopy** flag.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AddEx** method to create a document based on the "BASICD_U.VST" template that uses the default measurement system units.


```vb
Public Sub AddEx_Example() 
 
 Application.Documents.AddEx "BASICD_U.VST", visMSDefault 
 
End Sub
```


