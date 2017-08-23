---
title: "Объект мастера (издатель)"
keywords: vbapb10.chm1507327
f1_keywords: vbapb10.chm1507327
ms.prod: publisher
api_name: Publisher.Wizard
ms.assetid: c0a64ee9-d1fa-6dc7-5221-ff2d32874ea0
ms.date: 06/08/2017
ms.openlocfilehash: 9e1835832087f3db8f414636b0445f31baeb471e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizard-object-publisher"></a>Объект мастера (издатель)

Представляет макет публикации, связанных с публикацией или мастера, связанного с объектом макетов.
 


## <a name="example"></a>Пример

Свойство **[Мастер](document-wizard-property-publisher.md)** объект **документа**, **страницы**, **форму** и **ShapeRange** возвращает объект **мастера** . Следующий пример отчетов по публикации проекта, связанного с активной публикации, отображение его имя и текущие настройки.
 

 

```
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 MsgBox "Publication Design associated with " _ 
 &amp; "current publication: " _ 
 &amp; .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 MsgBox " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Примечание**  В зависимости от языковой версии Microsoft Publisher, используемая может появиться ошибка при использовании выше кода. В этом случае необходимо создать в обработчики ошибок для обхода ошибок. В следующем примере действует как приведенный выше код, но есть обработчики ошибок, созданных этой ситуации.
 


```
Sub ExampleWithErrorHandlers() 
 Dim wizTemp As Wizard 
 Dim wizproTemp As WizardProperty 
 Dim wizproAll As WizardProperties 
 
 Set wizTemp = ActiveDocument.Wizard 
 
 With wizTemp 
 Set wizproAll = .Properties 
 Debug.Print "Publication Design associated with " _ 
 &amp; "current publication: " _ 
 &amp; .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 If wizproTemp.Name = "Layout" Or wizproTemp _ 
 .Name = "Layout (Intl)" Then 
 On Error GoTo Handler 
 MsgBox " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 
Handler: 
 If Err.Number = 70 Then Resume Next 
 Else 
 MsgBox " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End If 
 End With 
 Next wizproTemp 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[SetId](wizard-setid-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](wizard-application-property-publisher.md)|
|[ID](wizard-id-property-publisher.md)|
|[Name](wizard-name-property-publisher.md)|
|[Родительский раздел](wizard-parent-property-publisher.md)|
|[Properties](wizard-properties-property-publisher.md)|

