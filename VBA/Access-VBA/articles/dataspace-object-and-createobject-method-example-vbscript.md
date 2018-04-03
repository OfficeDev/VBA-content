---
title: DataSpace Object and CreateObject Method Example (VBScript)
ms.prod: access
ms.assetid: 0132996b-e3e1-4dbe-95dc-4313824669b7
ms.date: 06/08/2017
---


# DataSpace Object and CreateObject Method Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016

The following example shows how to use the [CreateObject](http://msdn.microsoft.com/library/130debe5-31cf-4ab0-5f78-9adaec7d7126%28Office.15%29.aspx) method of the[RDS.DataSpace](http://msdn.microsoft.com/library/7db181d5-422b-49fe-b6af-a20f5da520ff%28Office.15%29.aspx) with the default business object,[RDSServer.DataFactory](http://msdn.microsoft.com/library/1de76cdd-34dc-8547-29aa-48ad6067bdea%28Office.15%29.aspx). To test this example, cut and paste this code between the <Body> and </Body> tags in a normal HTML document and name it  **DataSpaceVBS.asp**. ASP script will identify your server.




```vb

<!-- BeginDataSpaceVBS --><html>
<head><!--use the following META tag instead of adovbs.inc-->
<!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" --><meta name="VI60_DefaultClientScript" content=VBScript>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0"><title>DataSpace Object and CreateObject Method Example (VBScript)</title>
<style><!--
body {font-family: 'Verdana','Arial','Helvetica',sans-serif;
BACKGROUND-COLOR:white;COLOR:black;
}.thead {
background-color: #008080;font-family: 'Verdana','Arial','Helvetica',sans-serif;
font-size: x-small;color: white;
}.thead2 {
background-color: #800000;font-family: 'Verdana','Arial','Helvetica',sans-serif;
font-size: x-small;color: white;
}.tbody {
text-align: center;background-color: #f7efde;
font-family: 'Verdana','Arial','Helvetica',sans-serif;font-size: x-small;
}-->
</style></head> 
<body><h1>DataSpace Object and CreateObject Method Example (VBScript)</h1> 
<H2>RDS API Code Examples</H2><HR>
<H3>Using Query Method of RDSServer.DataFactory</H3> 
<!-- RDS.DataSpace ID rdsDS--><OBJECT ID="rdsDS" WIDTH=1 HEIGHT=1
CLASSID="CLSID:BD96C556-65A3-11D0-983A-00C04FC29E36"></OBJECT> 
<!-- RDS.DataControl with parameters set at run time --><OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33"
ID=RDS WIDTH=1 HEIGHT=1></OBJECT> 
<TABLE DATASRC=#RDS><TBODY>
<TR><TD><SPAN DATAFLD="FirstName"></SPAN></TD>
<TD><SPAN DATAFLD="LastName"></SPAN></TD></TR>
</TBODY></TABLE> 
<HR><INPUT TYPE=BUTTON NAME="Run" VALUE="Run"> 
<H4>Click Run -The <i>CreateObject</i> Method of the RDS.DataSpace Object Creates an instance of the RDSServer.DataFactory.
The <i>Query</i> Method of the RDSServer.DataFactory is used to bring back a Recordset.</H4> 
<Script Language="VBScript"> 
Dim rdsDFDim strServer
Dim strCnxnDim strSQL 
strServer = "http://<%=Request.ServerVariables("SERVER_NAME")%>"strCnxn = "Provider='sqloledb';Data Source=" &; _
"<%=Request.ServerVariables("SERVER_NAME")%>" &; ";" &; _"Integrated Security='SSPI';Initial Catalog='Northwind';"
strSQL = "Select FirstName, LastName from Employees" 
Sub Run_OnClick() 
Dim rs' Create Data Factory
Set rdsDF = rdsDS.CreateObject("RDSServer.DataFactory", strServer)'Get Recordset
Set rs = rdsDF.Query(strCnxn, strSQL)' Use RDS.DataControl to bind Recordset to data-aware Table above
RDS.SourceRecordset = rs 
End Sub</Script> 
</body></html>
<!-- EndDataSpaceVBS -->
```

The following example shows how to use the  **CreateObject** method to create an instance of a custom business object, VbBusObj.VbBusObjCls. It also uses the Active Server Pages scripting to identify the Web server name.
To see the complete example, open the sample applications selector. In the  **Client Tier** column, select **VBScript in Internet Explorer**. In the **Middle Tier** column, select **Custom Visual Basic Business Object**.



```vb
 
Sub Window_OnLoad() 
   strServer = "http://<%=Request.ServerVariables("SERVER_NAME")%>" 
   Set BO = ADS1.CreateObject("VbBusObj.VbBusObjCls", strServer) 
   txtConnect.Value = "dsn=Pubs;uid=MyUserId;pwd=MyPassword;" 
   txtGetRecordset.Value = "Select * From authors for Browse" 
End Sub 

```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

