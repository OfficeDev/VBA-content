---
title: Project.ReadWssData Method (Project)
keywords: vbapj.chm132840
f1_keywords:
- vbapj.chm132840
ms.prod: project-server
api_name:
- Project.Project.ReadWssData
ms.assetid: 97ff4d8e-8f0b-3b7f-9515-56376967e5bd
ms.date: 06/08/2017
---


# Project.ReadWssData Method (Project)

Returns the Project Workspace URLs for the active enterprise project as an XML string.


## Syntax

 _expression_. **ReadWssData**( ** _ProjectGuid_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProjectGuid_|Required|**String**|A valid Project GUID.|

### Return Value

 **Variant**


## Example

The following is an example of the XML schema definition.


```XML
<?xml version="1.0" encoding="utf-8"?>
<xs:schema xmlns:mstns="http://schemas.microsoft.com/office/project/server/webservices/ProjectWSSInfoDataSet/" 
           xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" 
           xmlns="http://schemas.microsoft.com/office/project/server/webservices/ProjectWSSInfoDataSet/" 
           attributeFormDefault="qualified" elementFormDefault="qualified" 
           targetNamespace="http://schemas.microsoft.com/office/project/server/webservices/ProjectWSSInfoDataSet/" 
           id="ProjectWSSInfoDataSet" xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xs:element msdata:IsDataSet="true" msdata:UseCurrentLocale="true" name="ProjectWSSInfoDataSet">
    <xs:complexType>
      <xs:choice minOccurs="0" maxOccurs="unbounded">
        <xs:element name="ProjWssInfo">
          <xs:complexType>
            <xs:sequence>
              <xs:element msdata:DataType="System.Guid, mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
                          name="WSS_SERVER_UID" type="xs:string" />
              <xs:element msdata:DataType="System.Guid, mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
                          minOccurs="0" name="PROJECT_UID" type="xs:string" />
              <xs:element minOccurs="0" name="PROJECT_WORKSPACE_URL" type="xs:string" />
              <xs:element minOccurs="0" name="PROJECT_ISSUES_URL" type="xs:string" />
              <xs:element minOccurs="0" name="PROJECT_RISKS_URL" type="xs:string" />
              <xs:element minOccurs="0" name="PROJECT_DOCUMENTS_URL" type="xs:string" />
              <xs:element minOccurs="0" name="PROJECT_WORKSPACE_NAME" type="xs:string" />
              <xs:element minOccurs="0" name="PROJECT_WORKSPACE_VSERVER_URL" type="xs:string" />
              <xs:element minOccurs="0" name="PROJECT_NAME" type="xs:string" />
              <xs:element minOccurs="0" name="PROJECT_COMMITMENTS_URL" type="xs:string" />
              <xs:element minOccurs="0" name="WSS_PWA_ADMIN_ROLE_ID" type="xs:int" />
              <xs:element minOccurs="0" name="WSS_PWA_PROJECT_MANAGER_ROLE_ID" type="xs:int" />
              <xs:element minOccurs="0" name="WSS_PWA_TEAM_MEMBER_ROLE_ID" type="xs:int" />
              <xs:element minOccurs="0" name="WSS_PWA_READER_ROLE_ID" type="xs:int" />
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:choice>
    </xs:complexType>
    <xs:unique msdata:PrimaryKey="true" name="Constraint1">
      <xs:selector xpath=".//mstns:ProjWssInfo" />
      <xs:field xpath="mstns:WSS_SERVER_UID" />
    </xs:unique>
  </xs:element>
</xs:schema>

```


