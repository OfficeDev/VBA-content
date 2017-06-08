---
title: Document.Clean Method (Visio)
keywords: vis_sdr.chm10552040
f1_keywords:
- vis_sdr.chm10552040
ms.prod: visio
api_name:
- Visio.Document.Clean
ms.assetid: 5fd5c6a6-1914-b29d-c0ae-0e5374d13a8e
ms.date: 06/08/2017
---


# Document.Clean Method (Visio)

Examines, reports, and repairs selected conditions in a document.


## Syntax

 _expression_ . **Clean**( **_nTargets_** , **_nActions_** , **_nAlerts_** , **_nFixes_** , **_bStopOnError_** , **_bLogFileName_** , **_nReserved_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _nTargets_|Optional| **Variant**|Indicates which parts of the document to examine. See Remarks for possible values.|
| _nActions_|Optional| **Variant**|Indicates which conditions to detect. See Remarks for possible values.|
| _nAlerts_|Optional| **Variant**|Indicates which detected conditions to report. See Remarks for possible values.|
| _nFixes_|Optional| **Variant**|Indicates which detected conditions to fix. See Remarks for possible values.|
| _bStopOnError_|Optional| **Variant**|Non-zero ( **True** ) to cause processing to stop if an error is encountered while attempting to fix a detected condition; zero ( **False** ) to allow processing to continue.|
| _bLogFileName_|Optional| **Variant**|Reserved for future use.|
| _nReserved_|Optional| **Variant**|Reserved for future use.|

### Return Value

Nothing


## Remarks

Internal Microsoft Visio developers use the  **Clean** method to validate and optimize the documents provided with Visio; third-party developers can use this method on their own documents.

It is suggested that developers use default values for  _nTargets_ , _nActions_ , _nAlerts_ , and _nFixes_ , and make a backup copy of a document before it is cleaned.

You can identify document changes made by the  **Clean** method by comparing saved VDX (XML) versions of the document, one version saved before the **Clean** method executes, and the other after.

The  _nTargets_ argument can be any combination of the values of the constants defined in **VisDocCleanTargets** in the Visio type library, and described in the following table.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDocCleanTargAll**|&;HFF|Examine all objects (default for  _nTargets_ ).|
| **visDocCleanTargFPages**|&;H1|Examine foreground pages.|
| **visDocCleanTargBPages**|&;H2|Examine background pages.|
| **visDocCleanTargMasters**|&;H4|Examine masters.|
| **visDocCleanTargStyles**|&;H8|Examine styles. |
| **visDocCleanTargDoc**|&;H10 |Examine document sheet.|
| **visDocCleanTargPageSheet**|&;H100|Examine page sheet(s). |
The nActions, nAlerts, and nFixes arguments can be any combination of the values of the constants defined in  **VisDocCleanActions** in the Visio type library, and described in the following table.



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visDocCleanActLocalFormulas**| &;H1| Detect unnecessary local overrides.|
| **visDocCleanActEmptyRowsAndSects**| &;H2| Detect empty local rows and sections.|
| **visDocCleanActNonDefaultFonts**| &;H4| Detect non-default font settings.|
| **visDocCleanActStaleResults**| &;H8| Detect results that don't match formulas.|
| **visDocCleanActMissingSubs**| &;H10| Detect missing subscriptions (cell dependencies).|
| **visDocCleanActConstantFormulas**| &;H20| Detect formulas that can be generated from the result.|
| **visDocCleanActNearZero**| &;H40| Detect results that are almost zero and change them to zero.|
| **visDocCleanActDuplicateSubs**| &;H80| Detect duplicate subscriptions (cell dependencies).|
| **visDocCleanActBadDisplayLists**| &;H100| Detect invalid display list linkages.|
| **visDocCleanActDeletedFields**| &;H400| Detect deleted fields.|
| **visDocCleanActBadFieldFormulas**| &;H800| Detect fields with missing or nonstandard formulas.|
| **visDocCleanActBadFieldMarks**| &;H1000| Detect fields with out-of-sync count and marker values. Change the position of escape characters to match character counts.|
| **visDocCleanActBadReferences**| &;H2000| Detect formulas with #Ref() errors.|
| **visDocCleanActAll**| &;H3FFF| Perform all actions.|
| **visDocCleanActDefault**| &;H1FD8| Default conditions to detect (default value of _nActions_ ).|
| **visDocCleanAlertDefault**| &;H0| Default conditions to report (default value of _nAlerts_ ).|
| **visDocCleanFixDefault**| &;H3D8| Default conditions to fix (default value for _nFixes_ ).|

## Example

The following procedure demonstrates one use of the  **Clean** method. In this case, the line pattern of a rectangle is overridden with the same value as it originally inherited, which creates an unnecessary local override. The **Clean** method is then executed, which detects the condition and posts an alert allowing the user to choose whether to fix the condition or not.


1. Create a new blank drawing.
    
2. Use the  **Rectangle** tool to draw a rectangle on the drawing page. If you view the shape in the ShapeSheet window, you can see that the color of the value ("1") in the LinePattern cell is black, indicating that the value is inherited.
    
3. Right-click the shape, point to  **Format**, click  **Line**, and in the  **Line** dialog box, reapply the same line pattern. This action creates a local value in the shape, or a local override. Now if you view the shape in the ShapeSheet window, you can see that the color of the value in the LinePattern cell is blue, indicating that the value is local.
    
4. Insert the  **Clean_Example** procedure shown below into your document's Microsoft Visual Basic for Applications project:
    
5. Run the  **Clean_Example** procedure (on the **View** tab, click **Macros**; then, in the  **Macros** dialog box, in the list of macros, select **ThisDocument.Clean_Example**, and then click  **Run**).
    

```vb
 
    Public Sub Clean_Example() 
     
        ActiveDocument.Clean, visDocCleanActLocalFormulas, _  
           visDocCleanActLocalFormulas, visDocCleanActLocalFormulas 
     
End Sub
```

Alerts appear on the drawing page asking whether you want to remove the unneeded local override. If you click  **Yes** and then reopen the ShapeSheet window, you can see that the color of the value in the LinePattern cell is once again black, indicating that the inherited value has been restored.


