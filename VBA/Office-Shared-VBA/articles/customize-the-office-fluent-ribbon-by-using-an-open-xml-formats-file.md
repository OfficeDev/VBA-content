---
title: Customize the Office Fluent Ribbon by Using an Open XML Formats File
ms.prod: office
ms.assetid: 562d79a2-c1eb-126a-1567-ddd0253f5972
ms.date: 06/08/2017
---


# Customize the Office Fluent Ribbon by Using an Open XML Formats File

The ribbon component of the Microsoft Office Fluent user interface gives users a flexible way to work with Office applications. Ribbon Extensibility (RibbonX) uses a simple, text-based, declarative XML markup to create and customize the ribbon. 

The code example in this topic shows how to add custom components to the ribbon for a single document, as opposed to adding application-level customizations. In the following steps, you add a custom tab, a custom group, and a custom button to the existing ribbon in Word. You also implement a callback procedure for the button that inserts a company name into the document. 

1. Create the customization file in any text editor and save the file with the name  **customUI.xml**.
    
2. Add the following XML markup to the file and then close and save the file. 
    
  ```XML
  <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"> 
  <ribbon> 
    <tabs> 
      <tab id="CustomTab" label="My Tab"> 
        <group id="SampleGroup" label="Sample Group"> 
          <button id="Button" label="Insert Company Name" size="large" onAction="ThisDocument.InsertCompanyName" /> 
        </group > 
      </tab> 
    </tabs> 
  </ribbon> 
</customUI> 

  ```

3. Create a folder on your desktop named  **customUI** and copy the XML customization file to the folder.
    
4. Validate the XML markup with a custom schema. 
    
     **Note**  This step is optional.
5. Create a document in Word 2007 and save it with the name  **RibbonSample.docm**.
    
6.  Open the Visual Basic Editor, add the following procedure to the **ThisDocument** code module, and then close and save the document.
    
  ```
  Sub InsertCompanyName(ByVal control As IRibbonControl) 
   ' Inserts the specified text at the beginning of a range or selection. 
   Dim MyText As String 
   Dim MyRange As Object 
   Set MyRange = ActiveDocument.Range 
   MyText = "Microsoft Corporation" 
   ' Range Example: Inserts text at the beginning 
   ' of the active document 
   MyRange.InsertBefore (MyText) 
   ' Selection Example: 
   'Selection.InsertBefore (MyText) 
End Sub 

  ```

7. Add a  **.zip** extension to the document file name and then double-click it to open the file.
    
8. Add the customization file to the container by dragging the customUI folder from the desktop to the .zip file.
    
9. Extract the  **.rels** file to your desktop. A **_rels** folder that contains the .rels file is copied to your desktop.
    
10. Open the  **.rels** file and add the following line between the last **Relationship** tag and the **Relationships** tag. This creates a relationship between the document file and the customization file.
    
     `<Relationship Id="someID" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml" />`
    
11. Close and save the file.
    
12. Add the _rels folder back to the container file by dragging it from the desktop, overwriting the existing file.
    
13. Rename the document file to its original name by removing the .zip extension.
    
14. Open the document and notice that the ribbon now displays the  **My Tab** tab.
    
15. Click the tab and notice the  **Sample Group** group with a button control.
    
16. Click the button to insert the company name into the document.
    


