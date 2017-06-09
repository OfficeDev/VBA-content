---
title: Connect an Item in the Data Store to a SQL Server Database
ms.prod: word
ms.assetid: 5c3ecc43-492d-0668-18f6-752b03dd2a54
ms.date: 06/08/2017
---


# Connect an Item in the Data Store to a SQL Server Database

Word enables you to generate documents by creating data-driven solutions. You can create a template document that includes a custom XML part and use content controls to bind to custom XML data by using XML mapping. Although the term  _template_ is used in this context, this document is not a Word template, but shares some characteristics of a Word template document. Then you can create a managed web-based application to build a new document based on the template document. The managed web-based application opens the template document, retrieves data from a Microsoft SQL Server database to build a new custom XML part, replaces the template document's custom XML part with the new part, and saves the template document as a new Word document.

This walkthrough explains how to build a new template document and how to create a server-side application that generates documents that display data that is stored in a Microsoft SQL Server database. To build this application, you will complete the following two tasks:

1. Create a Word template document.
    
2. Create a server-side web-based application that pulls data from a Microsoft SQL Server database and generates new documents based on the Word template document.
    
The programmatic objects that are used in this sample are as follows:

-  **[ContentControl](contentcontrol-object-word.md)**
    
-  **[ContentControls](contentcontrols-object-word.md)**
    
-  **[CustomXMLPart](http://msdn.microsoft.com/library/a4f90bac-01d6-bba4-f64b-a64e2b122cfd%28Office.15%29.aspx)** (in the Microsoft Office system core object model)
    
-  **[CustomXMLParts](http://msdn.microsoft.com/library/98c1c58e-a08d-6304-8626-1e6705917da3%28Office.15%29.aspx)** (in the Microsoft Office system core object model)
    
-  **[XMLMapping](xmlmapping-object-word.md)**
    
For more information about content controls, see  [Working with Content Controls](working-with-content-controls.md).

## Business Scenario: Create a Customer Document Generator

To create a Word document generator that connects an item in the data store to a Microsoft SQL Server database, you first build a template customer letter-generator document that contains content controls that map to an XML file. Next, you create a document-generation web-based application that enables you to select a company name to generate a custom document. The application retrieves customer data from a Microsoft SQL Server database and uses the customer letter generator to build a new document that displays customer data based on a user selection. The document displays the following information:


- Company Name
    
- Contact Name
    
- Contact Title
    
- Phone Number
    
Use the following general steps to create a Word document generator.


### To create a custom document generator and define the XML mappings for each content control


1. Open Word and create a blank document.
    
2. Add plain-text content controls to the document to bind to nodes in the data store.
    
    Content controls are predefined pieces of content. Word offers several kinds of content controls. This includes text blocks, check boxes, drop-down menus, combo boxes, calendar controls, and pictures. You can map these content controls to an element in an XML file. By using  [XPath](http://www.w3.org/TR/xpath) expressions, you can programmatically map content in an XML file to a content control. This enables you to write a simple and short application to manipulate and modify data in a document.
    
    To add a content control, on the  **Developer** tab, in the **Controls** group, click **Plain Text Content Control**.
    
    Add four plain-text content controls to the document. After you add each control, assign each one a title: Click the control; in the  **Controls** group, click **Properties**; in the  **Title** box, type a title for the control, as shown in the following list; and then click **OK**.
    
      - Company Name
    
  - Contact Name
    
  - Contact Title
    
  - Phone Number
    

    
    
    You can also use the following Visual Basic for Applications (VBA) code to add content controls to the document. Press ALT+F11 to open the Visual Basic editor, paste the code into the code window, click anywhere in the procedure, and then press F5 to run the code and add four content controls to your template document.
    


```vb
  Sub AddContentControls()

    Selection.Range.ContentControls.Add (wdContentControlText)
    Selection.ParentContentControl.Title = "Company Name"
    Selection.ParentContentControl.Tag = "Company Name"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph
    
    Selection.Range.ContentControls.Add (wdContentControlText)
    Selection.ParentContentControl.Title = "Contact Name"
    Selection.ParentContentControl.Tag = "Contact Name"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph
    
    Selection.Range.ContentControls.Add (wdContentControlText)
    Selection.ParentContentControl.Title = "Contact Title"
    Selection.ParentContentControl.Tag = "Contact Title"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph
    
    Selection.Range.ContentControls.Add (wdContentControlText)
    Selection.ParentContentControl.Title = "Phone Number"
    Selection.ParentContentControl.Tag = "Phone Number"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph
    
End Sub
```

3. Set the XML mapping on the content controls.
    
    XML mapping is a feature of Word that enables you to create a link between a document and an XML file. This creates true data/view separation between the document formatting and the custom XML data. 
    
    To load a custom XML part, you must first add a new data store to a  **[Document](document-object-word.md)** object by using the ** [Add](http://msdn.microsoft.com/library/f2c1588b-c11b-49ca-5db6-4fa4c26d10c5%28Office.15%29.aspx)** method of the **CustomXMLParts** collection. This appends a new, empty data store to the document. Because it is empty, you cannot use it yet. Next, you must load a custom XML part from an XML file into the data store, by calling the ** [Load](http://msdn.microsoft.com/library/f4d50c05-15bd-ccce-6198-9d6be401b29b%28Office.15%29.aspx)** method of the **CustomXMLPart** object that uses a valid path to an XML file as the parameter.
    
4. Save the document, naming it CustomerLetterGenerator.docm.
    
	|**Note**|
	|:-----|  
	|Because it contains VBA code, you must save the document as a macro-enabled document file (.docm).|

The following procedure explains how to map a content control to a sample custom XML file. You create a valid custom XML file, save the file, and then you use Visual Basic for Applications (VBA) code to add to the template document a data store that contains the information to which you want to map.


### To set an XML mapping on a content control


1. Create a text file and save it as CustomerData.xml.
    
2. Copy the following XML code into the text file and save the file.
    
```XML
  <?xml version="1.0"?> 
<Customer> 
  <CompanyName>Alfreds Futterkiste</CompanyName> 
  <ContactName>Maria Anders</ContactName> 
  <ContactTitle>Sales Representative</ContactTitle> 
  <Phone>030-0074321</Phone> 
</Customer> 

```

3. Open .
    
4. Press ALT+F11 to open the Visual Basic editor, paste the following code into the code window, click anywhere in the procedure, and then press F5 to run the code and attach an XML file to your template document so that it becomes an available data store item.
    
```vb
  Public Sub LoadXML()

  ' Load CustomerData.xml file
   ActiveDocument.CustomXMLParts.Add
   ActiveDocument.CustomXMLParts(ActiveDocument.CustomXMLParts.Count).Load ("C:\CustomerData.xml")
End Sub
```


   	|**Note**|
	|:-----|  
	|There are at least three default custom XML parts that are always created with a document. These are 'Cover pages', 'Doc properties' and 'App properties'. In addition, various other custom XML parts may be created on a given computer, depending on several factors. These include which add-ons are installed, connections with SharePoint, and so on. Calling the  **Add** method on the **CustomXMLParts** collection in the previous code adds an additional XML part, into which the XML file is loaded. It is on that part that the **Load** method is called, in the next line of code. To determine the index number of the part into which to load the XML file, it is necessary to pass the count of custom XML parts to the **Load** method. `ActiveDocument.CustomXMLParts(ActiveDocument.CustomXMLParts.Count).Load ("C:\CustomerData.xml")`| 
    
     
5. Set an XML mapping on a content control that refers to a node in the added data store. To create an XML mapping, use an XPath expression that points to the node in the custom XML data part to which you want to map a content control. After you add a data store to your document (and the data store points to a valid XML file), you are ready to map one of its nodes to a content control. To do this, pass a  **String** that contains a valid **XPath** to a **ContentControl** object by using the **SetMapping** method of the **XMLMapping** object (by using the **XMLMapping** property of the **ContentControl** object). Open the Visual Basic editor and run the following VBA code to bind content controls to items in the data store.
    
```vb
  Public Sub MapXML()

    Dim strXPath1 As String
    strXPath1 = "/Customer/CompanyName"
    ActiveDocument.ContentControls(1).XMLMapping.SetMapping strXPath1
    
    Dim strXPath2 As String
    strXPath2 = "/Customer/ContactName"
    ActiveDocument.ContentControls(2).XMLMapping.SetMapping strXPath2
    
    Dim strXPath3 As String
    strXPath3 = "/Customer/ContactTitle"
    ActiveDocument.ContentControls(3).XMLMapping.SetMapping strXPath3
    
    Dim strXPath4 As String
    strXPath4 = "/Customer/Phone"
    ActiveDocument.ContentControls(4).XMLMapping.SetMapping strXPath4

```


## Create a Server-Side Application That Pulls Data from a SQL Server Database and Generates a New Document

You can create a Web-based application that enables users to select a company name and generate a custom letter. The Web-based application retrieves customer data from a SQL Server database, opens the customer letter template document, and creates a new document that displays customer data based on a user selection. This Web-based application does not require the use of Word or VBA. You can use your favorite managed code (Visual Basic .NET or C#) language to build this application.

|**Note**|
|:-----|  
|The Web-based application shown here gets its data from the Northwind.mdf database. This database was installed with previous versions of SQL Server and Office. If you do not have the Northwind database on your computer, you can download it from the following site:  [http://code.msdn.microsoft.com/northwind/Release/ProjectReleases.aspx?ReleaseId=1401](http://code.msdn.microsoft.com/northwind/Release/ProjectReleases.aspx?ReleaseId=1401)|

To build this application, do the following:


### To create a server-side application that pulls data from a SQL Server database and generates a new document


1. Open Visual Studio or Visual Web Developer.
    
2. Create an ASP.NET Web application and name it SqlServerSample.
    
    In the following steps, you'll connect the ASP.NET Web application to a SQL Server database.
    
3. Add the following connection string to the Web.config file in your Visual Studio project.
    
```XML
  <connectionStrings>
 <add name="NorthwindConnectionString"
     connectionString="data source=(local);database=Northwind; integrated security=SSPI;persist security info=false;"
     providerName="System.Data.SqlClient" />
</connectionStrings>
```

4. In your Visual Studio project, add the CustomerLetterGenerator.docm document to the  **App_Data** folder: Right-click **App_Data**, point to  **Add**, click  **Existing Item**, browse to the location where you saved the document, select it, and then click  **Add**.
    
5. Add a reference to WindowsBase.dll to your project: Right-click  **References**, click  **Add Reference**, click the  **.NET** tab, select **WindowsBase**, and then click  **OK**.
    
6. Download and install the  [Microsoft .NET Framework 4.0](http://www.microsoft.com/downloads/details.aspx?FamilyID=9cfb2d51-5ff4-4491-b0e5-b386f32c0992&;displaylang=en)
    
7. Configure the assembly in the Web.config file as follows.
    
```XML
  <compilation debug="false">
  <assemblies>
    <add assembly="WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35"/>
   </assemblies>
 </compilation>
```

8. Add a Web Form to your project: On the  **Project** menu, click **Add New Item**; under  **Installed Templates**, click  **Web**; select  **Web Form**, and then click  **Add**.
    
9. In the Solution Explorer, right-click  **Properties**, and then click  **Open**.
    
10. On the  **Web** tab, under **Start Action**, select  **Specific Page**, click the browse button, and navigate to the page  **WebForm1.aspx**.
    
11. Add the following code to the  **WebForm1.aspx** file, overwriting the part of the file bounded by the opening and closing <html> tags.
    
```HTML
  <html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
     <title>Data-Driven Document Generation - SQL Server Sample</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <h1>Customer Letter Generator</h1>
            <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; height: 12%">
                <tr>
                    <td>
                        Choose a customer:</td>
                    <td>
                        <asp:DropDownList 
                           ID="ddlCustomer"
                           runat="server"
                           AutoPostBack="True"
                           DataSourceID="CustomerData"
                           DataTextField="CompanyName"
                           DataValueField="CustomerID" 
                           Width="301px">
                        </asp:DropDownList>
                        <asp:SqlDataSource
                          ID="CustomerData"
                          runat="server"
                          ConnectionString="<%$ ConnectionStrings:NorthwindConnectionString %>"
                          SelectCommand="SELECT [CustomerID], [CompanyName] FROM [Customers]" ProviderName="<%$ ConnectionStrings:NorthwindConnectionString.ProviderName %>">
                        </asp:SqlDataSource>
                    </td>
                </tr>
          </table>
        </div>
        <br />
        <asp:Button
          ID="Button1"
          runat="server"
          OnClick="SubmitBtn_Click" 
          Text="Create Letter"
          Width="123px" />    
    </form>
</body>
</html>

```

12. Depending on the coding language you use, add the following Visual Basic .NET or C# code to the appropriate  **WebForm1.aspx** code-behind page in your project.
    

## Sample Code: Visual Basic .NET

The following Visual Basic .NET sample shows how to bind to a SQL Server database to retrieve data based on a customer selection and create a new document based on the CustomerLetterGenerator.docm template document. Add the following code to the  **WebForm1.aspx.vb** file, overwriting the existing code in the file.


```VB.net
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.IO.Packaging
Imports System.Linq
Imports System.Xml
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls

Public Class WebForm1

    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Const strRelRoot As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

    Private Sub CreateDocument()
        ' Get the template document file and create a stream from it
        Const DocumentFile As String = "~/App_Data/CustomerLetterGenerator.docm"

        ' Read the file into memory
        Dim buffer() As Byte = File.ReadAllBytes(Server.MapPath(DocumentFile))
        Dim memoryStream As MemoryStream = New MemoryStream(buffer, True)
        buffer = Nothing

        ' Open the document in the stream and replace the custom XML part
        Dim pkgFile As Package = Package.Open(memoryStream, FileMode.Open, FileAccess.ReadWrite)
        Dim pkgrcOfficeDocument As PackageRelationshipCollection = pkgFile.GetRelationshipsByType(strRelRoot)
        For Each pkgr As PackageRelationship In pkgrcOfficeDocument
            If (pkgr.SourceUri.OriginalString = "/") Then

                ' Get the root part
                Dim pkgpRoot As PackagePart = pkgFile.GetPart(New Uri(("/" + pkgr.TargetUri.ToString), UriKind.Relative))

                ' Add a custom XML part to the package
                Dim uriData As Uri = New Uri("/customXML/item1.xml", UriKind.Relative)
                If pkgFile.PartExists(uriData) Then

                    ' Delete part "/customXML/item1.xml" part
                    pkgFile.DeletePart(uriData)
                End If

                ' Load the custom XML data
                Dim pkgprtData As PackagePart = pkgFile.CreatePart(uriData, "application/xml")
                GetDataFromSQLServer(pkgprtData.GetStream, ddlCustomer.SelectedValue)
            End If
        Next

        ' Close the file
        pkgFile.Close()

        ' Return the result
        Response.ClearContent()
        Response.ClearHeaders()
        Response.AddHeader("content-disposition", "attachment; filename=document.docx")
        Response.ContentEncoding = System.Text.Encoding.UTF8
        memoryStream.WriteTo(Response.OutputStream)
        memoryStream.Close()
        Response.End()
    End Sub

    Private Sub GetDataFromSQLServer(ByVal stream As Stream, ByVal customerID As String)

        'Connect to a SQL Server database and get data
        Dim source As String = ConfigurationManager.ConnectionStrings("NorthwindConnectionString").ConnectionString
        Const SqlStatement As String = "SELECT CompanyName, ContactName, ContactTitle, Phone FROM Customers WHERE CustomerID=@customerID"
        Dim conn As SqlConnection = New SqlConnection(source)
        conn.Open()
        Dim cmd As SqlCommand = New SqlCommand(SqlStatement, conn)
        cmd.Parameters.AddWithValue("@customerID", customerID)
        Dim dr As SqlDataReader = cmd.ExecuteReader
        If dr.Read Then
            Dim writer As XmlWriter = XmlWriter.Create(stream)
            writer.WriteStartElement("Customer")
            writer.WriteElementString("CompanyName", CType(dr("CompanyName"), String))
            writer.WriteElementString("ContactName", CType(dr("ContactName"), String))
            writer.WriteElementString("ContactTitle", CType(dr("ContactTitle"), String))
            writer.WriteElementString("Phone", CType(dr("Phone"), String))
            writer.WriteEndElement()
            writer.Close()
        End If
        dr.Close()
        conn.Close()
    End Sub

    Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As EventArgs)
        CreateDocument()
    End Sub

End Class
```


## Sample Code: C#

The following C# sample shows how to bind to a SQL Server database to retrieve data based on a customer selection and create a new document based on the CustomerLetterGenerator.docm template document. Add the following code to the  **WebForm1.Aspx.cs** file, copying over the existing code.


```C#
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SQLServerSample
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        private const string strRelRoot = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

        private void CreateDocument()
        {
            // Get the template document file and create a stream from it
            const string DocumentFile = @"~/App_Data/CustomerLetterGenerator.docm";
            
            // Read the file into memory
            byte[] buffer = File.ReadAllBytes(Server.MapPath(DocumentFile));
            MemoryStream memoryStream = new MemoryStream(buffer, true);
            buffer = null;

            // Open the document in the stream and replace the custom XML part
            Package pkgFile = Package.Open(memoryStream, FileMode.Open, FileAccess.ReadWrite);
            PackageRelationshipCollection pkgrcOfficeDocument = pkgFile.GetRelationshipsByType(strRelRoot);
            foreach (PackageRelationship pkgr in pkgrcOfficeDocument)
            {
                if (pkgr.SourceUri.OriginalString == "/")
                {
                    // Get the root part
                    PackagePart pkgpRoot = pkgFile.GetPart(new Uri("/" + pkgr.TargetUri.ToString(), UriKind.Relative));

                    // Add a custom XML part to the package
                    Uri uriData = new Uri("/customXML/item1.xml", UriKind.Relative);

                    if (pkgFile.PartExists(uriData))
                    {
                        // Delete document "/customXML/item1.xml" part
                        pkgFile.DeletePart(uriData);
                    }
                    // Load the custom XML data
                    PackagePart pkgprtData = pkgFile.CreatePart(uriData, "application/xml");
                    GetDataFromSQLServer(pkgprtData.GetStream(), ddlCustomer.SelectedValue);
                }
            }

            // Close the file
            pkgFile.Close();

            // Return the result
            Response.ClearContent();
            Response.ClearHeaders();
            Response.AddHeader("content-disposition", "attachment; filename=CustomLetter.docx");
            Response.ContentEncoding = System.Text.Encoding.UTF8;

            memoryStream.WriteTo(Response.OutputStream);

            memoryStream.Close();

            Response.End();
        }

        private void GetDataFromSQLServer(Stream stream, string customerID)
        {
            // Connect to a SQL Server database and get data
            String source = System.Configuration.ConfigurationManager.ConnectionStrings["NorthwindConnectionString"].ConnectionString;            
            const string SqlStatement =
                "SELECT CompanyName, ContactName, ContactTitle, Phone FROM Customers WHERE CustomerID=@customerID";

            using (SqlConnection conn = new SqlConnection(source))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(SqlStatement, conn);
                cmd.Parameters.AddWithValue("@customerID", customerID);
                SqlDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    XmlWriter writer = XmlWriter.Create(stream);
                    writer.WriteStartElement("Customer");
                    writer.WriteElementString("CompanyName", (string)dr["CompanyName"]);
                    writer.WriteElementString("ContactName", (string)dr["ContactName"]);
                    writer.WriteElementString("ContactTitle", (string)dr["ContactTitle"]);
                    writer.WriteElementString("Phone", (string)dr["Phone"]);
                    writer.WriteEndElement();
                    writer.Close();
                }
                dr.Close();
                conn.Close();
            }
        }

        protected void SubmitBtn_Click(object sender, EventArgs e)
        {
            CreateDocument();
        }
    }
}
```

For more information about working with ASP.NET 2.0, see  [http://www.asp.net/get-started/](http://www.asp.net/get-started).

This article explains how to extract data from a SQL Server database and insert it into your template document. You can also extract the data from other data sources, including, for example, Access and Excel. For more information about how to connect to data in those applications programmatically, see the Access and Excel developer documentation.


