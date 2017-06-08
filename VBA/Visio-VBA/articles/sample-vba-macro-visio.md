---
title: Sample VBA Macro (Visio)
keywords: vis_sdr.chm81901862
f1_keywords:
- vis_sdr.chm81901862
ms.prod: visio
ms.assetid: 70ffb571-5794-875c-869a-a68a5e1b8ac8
ms.date: 06/08/2017
---


# Sample VBA Macro (Visio)

For each drawing file that is open in the Visio instance, the sample Visual Basic for Applications (VBA) macro shown below does the following:


- Logs the name and path of the drawing file in the  **Immediate** window
    
- Logs the name of each page in the  **Immediate** window
    

```vb
Public Sub ShowNames()  
 
    'Declare object variables as Visio object types.  
    Dim vsoPage As Visio.Page  
    Dim vsoDocument As Visio.Document  
    Dim vsoDocuments As Visio.Documents  
    Dim vsoPages As Visio.Pages  
 
    'Iterate through all open documents.  
    Set vsoDocuments  = Application.Documents  
    For Each vsoDocument In vsoDocuments   
 
        'Print the drawing name in the Visual Basic Editor  
        'Immediate window.  
        Debug.Print vsoDocument.FullName  
 
        'Iterate through all pages in a drawing.  
        Set vsoPages = vsoDocument.Pages  
        For Each vsoPage In vsoPages 
  
            'Print the page name in the Visual Basic Editor  
            'Immediate window.  
            Debug.Print Tab(5); vsoPage.Name 
  
        Next  
 
    Next  
 
End Sub
```

|**Note**|
|:-----|  
|Here is an example of the program's output, assuming drawings named Office.vsd and Recycle.vsd are open and have been saved in the specified locations. The locations shown are not those in which Visio saves drawings by default.|


|**Sample output**|**Description**|
|:-----|:-----|
|```C:\documents\drawings\Office.vsd```| The name of the first drawing|
|```Background-1```|The name of page 1|
|```Background-2```|The name of page 2|
|```C:\documents\drawings\Recycle.vsd```|The name of the second drawing|
|```Page-1```|The name of page 1|
|```Page-2```|The name of page 2|
|```Page-3```|The name of page 3|

You can find more information about writing a program in the VBA environment and about the Visual Basic Editor in Visual Basic Help (in the Visual Basic Editor window, on the  **Help** menu, click **Microsoft Visual Basic Help**).
You can find details about using a specific Visio object, property, method, enumeration, or event in this reference.

