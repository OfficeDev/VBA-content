---
title: Project.GetWinprojURLs Method (Project)
keywords: vbapj.chm131098
f1_keywords:
- vbapj.chm131098
ms.prod: project-server
api_name:
- Project.Project.GetWinprojURLs
ms.assetid: 4ea8b044-9397-d17f-b057-d39105d83cf8
ms.date: 06/08/2017
---


# Project.GetWinprojURLs Method (Project)

Returns the various URLs associated with the active enterprise project as an XML string.


## Syntax

 _expression_. **GetWinprojURLs**

 _expression_ A variable that represents a **Project** object.


### Return Value

 **Variant**


## Example

The following is an example of the XML schema definition.


```XML
<?xml version="1.0" encoding="utf-8" ?>
<xs:schema id="WinprojURLsDataSet" 
           targetNamespace="http://schemas.microsoft.com/office/project/server/webservices/WinprojURLsDataSet/"
 xmlns:mstns="http://schemas.microsoft.com/office/project/server/webservices/WinprojURLsDataSet/" 
           xmlns="http://schemas.microsoft.com/office/project/server/webservices/WinprojURLsDataSet/"
 xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata"
 xmlns:NameSpace1="urn:schemas-microsoft-com:xml-msdatasource" attributeFormDefault="qualified"
 elementFormDefault="qualified">
  <xs:annotation>
    <xs:appinfo source="urn:schemas-microsoft-com:xml-msdatasource">
      <DataSource DefaultConnectionIndex="0" Modifier="AutoLayout, AnsiClass, NotPublic, Public" 
                  xmlns="urn:schemas-microsoft-com:xml-msdatasource">
        <Connections></Connections>
        <Tables></Tables>
        <Sources></Sources>
      </DataSource>
    </xs:appinfo>
  </xs:annotation>
  <xs:element name="WinprojURLsDataSet" msdata:IsDataSet="true">
    <xs:complexType>
      <xs:choice minOccurs="0" maxOccurs="unbounded">
        <xs:element name="WinprojURLs">
          <xs:complexType>
            <xs:sequence>
              <xs:element name="PROJECT_CENTER_URL" type="xs:string" minOccurs="0" />
              <xs:element name="RESOURCE_CENTER_URL" type="xs:string" minOccurs="0" />
              <xs:element name="PORTFOLIO_ANALYZER_URL" type="xs:string" minOccurs="0" />
              <xs:element name="GLOBAL_ISSUES_URL" type="xs:string" minOccurs="0" />
              <xs:element name="GLOBAL_RISKS_URL" type="xs:string" minOccurs="0" />
              <xs:element name="GLOBAL_DOCUMENTS_URL" type="xs:string" minOccurs="0" />
              <xs:element name="STATUS_REPORTS_URL" type="xs:string" minOccurs="0" />
              <xs:element name="APPROVALS_URL" type="xs:string" minOccurs="0" />
              <xs:element name="TIMESHEETS_URL" type="xs:string" minOccurs="0" />
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:choice>
    </xs:complexType>
  </xs:element>
</xs:schema>
<?xml version="1.0" encoding="utf-8" ?>
<xs:schema id="WinprojURLsDataSet" 
           targetNamespace="http://schemas.microsoft.com/office/project/server/webservices/WinprojURLsDataSet/"
 xmlns:mstns="http://schemas.microsoft.com/office/project/server/webservices/WinprojURLsDataSet/" 
           xmlns="http://schemas.microsoft.com/office/project/server/webservices/WinprojURLsDataSet/"
 xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata"
 xmlns:NameSpace1="urn:schemas-microsoft-com:xml-msdatasource" attributeFormDefault="qualified"
 elementFormDefault="qualified">
  <xs:annotation>
    <xs:appinfo source="urn:schemas-microsoft-com:xml-msdatasource">
      <DataSource DefaultConnectionIndex="0" Modifier="AutoLayout, AnsiClass, NotPublic, Public" 
                  xmlns="urn:schemas-microsoft-com:xml-msdatasource">
        <Connections></Connections>
        <Tables></Tables>
        <Sources></Sources>
      </DataSource>
    </xs:appinfo>
  </xs:annotation>
  <xs:element name="WinprojURLsDataSet" msdata:IsDataSet="true">
    <xs:complexType>
      <xs:choice minOccurs="0" maxOccurs="unbounded">
        <xs:element name="WinprojURLs">
          <xs:complexType>
            <xs:sequence>
              <xs:element name="PROJECT_CENTER_URL" type="xs:string" minOccurs="0" />
              <xs:element name="RESOURCE_CENTER_URL" type="xs:string" minOccurs="0" />
              <xs:element name="PORTFOLIO_ANALYZER_URL" type="xs:string" minOccurs="0" />
              <xs:element name="GLOBAL_ISSUES_URL" type="xs:string" minOccurs="0" />
              <xs:element name="GLOBAL_RISKS_URL" type="xs:string" minOccurs="0" />
              <xs:element name="GLOBAL_DOCUMENTS_URL" type="xs:string" minOccurs="0" />
              <xs:element name="STATUS_REPORTS_URL" type="xs:string" minOccurs="0" />
              <xs:element name="APPROVALS_URL" type="xs:string" minOccurs="0" />
              <xs:element name="TIMESHEETS_URL" type="xs:string" minOccurs="0" />
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:choice>
    </xs:complexType>
  </xs:element>
</xs:schema>
```


