
# CubeDef Example (VBScript)

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

This example displays cube metadata on a web page.




```
 
<%@ Language=VBScript %> 
<% 
Response.Buffer=True 
'Response.Expires=0 
%> 
<html> 
<head> 
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0"> 
</head> 
<body> 
 
<% 
Server.ScriptTimeout=360 
Dim cat,cdf,di,hi,le,mem,strServer,strSource,strCubeName 
 
'************************************************************************ 
'*** Set Session Variables 
'************************************************************************ 
Session("CubeName") = Request.Form("strCubeName") 
Session("CatalogName") = Request.Form("strCatalogName") 
Session("ServerName") = Request.Form("strServerName") 
Session("chkDim") = Request.Form("chkDimension") 
Session("chkHier") =  Request.Form("chkHierarchy") 
Session("chkLev") =  Request.Form("chkLevel") 
 
'************************************************************************ 
'*** Create Catalog Object 
'************************************************************************************ 
Set cat = Server.CreateObject("ADOMD.Catalog") 
 
If Len(Session("ServerName")) > 0 Then 
   cat.ActiveConnection = "Data Source='" &amp; Session("ServerName") &amp; "';Initial Catalog='" &amp; Session("CatalogName") &amp; "';Provider='msolap';" 
Else 
'************************************************************************************ 
'*** Must set OLAPServerName to OLAP Server that is 
'*** present on network 
'************************************************************************ 
OLAPServerName = "Please set to present OLAP Server" 
   cat.ActiveConnection = "Data Source=" &amp; OLAPServerName &amp; _ 
      ";Initial Catalog=FoodMart;Provider=msolap;" 
   Session("ServerName") = OLAPServerName 
   Session("InitialCatalog") = "FoodMart" 
End if 
 
If Len(Session("CubeName")) > 0 Then 
   Set cdf = cat.CubeDefs(Session("CubeName")) 
Else 
   Set cdf = cat.CubeDefs("Sales") 
   Session("CubeName")="Sales" 
End if 
 
'************************************************************************ 
'*** Collect Information in HTML Form 
'************************************************************************ 
%> 
<form action="ASPADOCubeDoc.asp" method="post" id="form1" name="form1"> 
<table> 
   <tr> 
      <td> 
      <b>Olap Server name:  </b><br><input type="text" id="strServerName" name="strServerName" value="<%=Session("ServerName")%>" size="20"><br> 
 
      <b>Catalog Name:  </b><br><input type="text" id="strCatalogName" name="strCatalogName" value="<%=Session("CatalogName")%>" size="20"><br> 
 
      <b>Cube Name:  </b><br><input type="text" id="strCubeName" name="strCubeName" value="<%=Session("CubeName")%>" size="20"> 
      </td> 
      <td <TD> 
         <b>Add Property Detail:  </b><br> 
         Dimension Detail: <input type="checkbox" id="chkDimension" name="chkDimension"><br> 
 
         Hierarchy Detail: <input type="checkbox" id="chkHierarchy" name="chkHierarchy"><br> 
 
         Level Detail: <input type="checkbox" id="chkLevel" name="chkLevel"> 
      </td>  
   </tr> 
</table> 
<input type="submit" value="Cube Information" id="submit1" name="submit1"><input type="reset" value="Reset" id="reset1" name="reset1"> 
</form> 
<% 
 
'************************************************************************ 
'*** Start of Report 
'************************************************************************ 
Response.Write "<H3>Report for " &amp; Session("CubeName") &amp; " Cube</H3>" 
Response.Write "<OL TYPE='i'>" 
 
'************************************************************************ 
'*** Show properties of Cube 
'************************************************************************ 
            For i = 0 To cdf.Properties.Count - 1 
               Response.Write "<LI>" 
               Response.Write "<FONT size=-2>" &amp; cdf.Properties(i).Name &amp; ": " &amp; cdf.Properties(i).Value &amp; "</FONT>" 
            Next 
            Response.Write "</OL>" 
            Response.Write "<UL TYPE='SQUARE'>"    
 '************************************************************************ 
'*** Loop to display Dimension Name and Properties if Check box is  
'*** Checked 
'************************************************************************ 
      For di = 0 To cdf.Dimensions.Count - 1 
         Response.Write "<LI>" 
         Response.Write "<FONT size=4><B>Dimension: " &amp; _ 
            cdf.Dimensions(di).Name &amp; "</B></FONT>" 
         If Request.Form("chkDimension") = "on" Then 
            Response.Write "<OL TYPE='1'>" 
            For i = 0 To cdf.Dimensions(di).Properties.Count - 1 
               Response.Write "<LI>" 
               Response.Write "<FONT size=-2>" &amp; _ 
                  cdf.Dimensions(di).Properties(i).Name &amp; ": " &amp; _ 
                  cdf.Dimensions(di).Properties(i).Value &amp; "</FONT>" 
            Next 
            Response.Write "</OL>" 
         End If 
         Response.Write "<UL TYPE= 'Circle'>" 
'************************************************************************ 
'*** Loop to display Hierarchy Name and Properties if Check box is  
'*** Checked 
'************************************************************************ 
         For hi = 0 To cdf.Dimensions(di).Hierarchies.Count - 1 
            Response.Write "<LI>" 
            Response.Write "<FONT size=3><B>Hierarchy: " &amp; _ 
               cdf.Dimensions(di).Hierarchies(hi).Name &amp; "</B></FONT>" 
            If Request.Form("chkHierarchy") = "on" Then 
               Response.Write "<OL TYPE='1'>" 
               For i = 0 To _ 
                  cdf.Dimensions(di).Hierarchies(hi).Properties.Count - 1 
                  Response.Write "<LI>" 
                  Response.Write "<FONT size=-2>" &amp; _ 
                     cdf.Dimensions(di).Hierarchies(hi).Properties(i)._ 
                     Name &amp; ": " &amp; _ 
                     cdf.Dimensions(di).Hierarchies(hi).Properties(i)._ 
                     Value &amp; "</FONT>" 
               Next 
               Response.Write "</OL>" 
            End If 
            Response.Write "<UL TYPE='Disc'>" 
'************************************************************************ 
'*** Loop to display Level Name and Properties if Check box is Checked 
'************************************************************************ 
      For le = 0 To cdf.Dimensions(di).Hierarchies(hi).Levels.Count - 1 
               Response.Write "<LI>" 
               Response.Write "<FONT size=2><B>Level: " &amp; _ 
                  cdf.Dimensions(di).Hierarchies(hi).Levels(le).Name &amp; _ 
                  " with a Member Count of: " &amp; _ 
                  cdf.Dimensions(di).Hierarchies(hi).Levels(le)._ 
                  Properties("LEVEL_CARDINALITY") &amp; "</B></FONT>" 
               If Request.Form("chkLevel") = "on" Then 
                  Response.Write "<OL TYPE='1'>" 
                  For i = 0 To  
                     cdf.Dimensions(di).Hierarchies(hi).Levels(le)._ 
                     Properties.Count - 1 
                     Response.Write "<LI>" 
                     Response.Write "<FONT size=-2>" &amp; _ 
                        cdf.Dimensions(di).Hierarchies(hi).Levels(le)._ 
                        Properties(i).Name &amp; ": " &amp; _ 
                        cdf.Dimensions(di).Hierarchies(hi).Levels(le)._ 
                        Properties(i).Value &amp; "</FONT>" 
                  Next 
                  Response.Write "</OL>" 
               End If 
            Next 
            Response.Write "</UL>" 
         Next 
         Response.Write "</UL>" 
      Next 
      Response.Write "</UL>" 
%> 
</body> 
</html> 

```

 **ACCESS SUPPORT RESOURCES**<br><br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br><br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br><br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&amp;tab=question&amp;status=all&amp;auth=1)<br><br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br><br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br><br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br><br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br><br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)
