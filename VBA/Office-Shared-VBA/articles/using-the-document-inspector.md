---
title: Using the Document Inspector
ms.prod: office
ms.assetid: 62180311-ee41-1812-797d-3b5814add284
ms.date: 06/08/2017
---


# Using the Document Inspector

The  **Document Inspector** gives users an easy way to examine documents for personal or sensitive information, text phrases, and other document contents. They can use the **Document Inspector** to remove unwanted information; for example, before distributing a document.


 **Note**  Microsoft does not support the automatic removal of hidden information for signed or protected documents, or for documents that use Information Rights Management (IRM). We recommend that you run the  **Document Inspector** before you sign a document or invoke IRM on a document.


As a developer, you can use the Document Inspector framework to extend the built-in modules and integrate your extensions into the standard user interface. 

The  **Document Inspector** in Microsoft Word, Microsoft Excel, and Microsoft PowerPoint includes the following enhancements.

## Built-in Document Inspector Modules

The  **Document Inspector** has modules that help users inspect and fix specific elements of a given document. The **Document Inspector** includes the following built-in modules:

For all Office documents:


- Embedded Documents
    
- OLE Objects and Packages
    
- Data Models
    
- Content Apps
    
- Task Pane Apps
    
- Macros and VBA Modules
    
- Legacy Macros (XLM and WordBasic)
    
For Excel documents:


- PivotTables and Slicers
    
- PivotCharts
    
- Cube Formulas
    
- Timelines (cache)
    
- Custom XML Data
    
- Comments and Annotations
    
- Document Properties and Personal Information
    
- Headers and Footers
    
- Hidden Rows and Columns
    
- Hidden Worksheets and Names
    
- Invisible Content
    
- External Links and Data Functions
    
- Excel Surveys
    
- Custom Worksheet Properties
    
For PowerPoint documents:


- Comments and Annotations
    
- Document Properties and Personal Information
    
- Invisible On-Slide Content
    
- Off-slide Content
    
- Presentation Notes
    
For Word documents:


- Comments, Revisions, Versions, and Annotations
    
- Document Properties and Personal Information. This includes metadata, Microsoft SharePoint properties, custom properties, and other content information.
    
- Custom XML Data
    
- Headers, Footers, and Watermarks
    
- Invisible Content
    
- Hidden Text
    

## Opening the Document Inspector

To open the  **Document Inspector**:


1. Click the  **File** tab, and then click **Info**.
    
2. Click  **Check for Issues**.
    
3. Click  **Inspect Document**.
    


Use the  **Document Inspector** dialog box to select the type or types of data to find in the document.

After the modules complete the inspection, the  **Document Inspector** displays the results for each module in a dialog box. If a given module finds data, the dialog box includes a **Remove All** button that you can click to remove that data. If the module does not find data, the dialog box displays a message to that effect.

If you choose to remove the data for a given module, the dialog box displays descriptive text that indicates whether the operation was successful or not. If the  **Document Inspector** encounters errors during the operation, the module is flagged, displays an error message, and the data for that module does not change.


