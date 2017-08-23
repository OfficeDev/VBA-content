---
title: "Объект WizardProperties (издатель)"
keywords: vbapb10.chm1572863
f1_keywords: vbapb10.chm1572863
ms.prod: publisher
api_name: Publisher.WizardProperties
ms.assetid: b3feecf2-ffbb-79de-8586-6a64df1b816a
ms.date: 06/08/2017
ms.openlocfilehash: b6c73e1993d361e9b935de9b93ec69a331df01dd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardproperties-object-publisher"></a>Объект WizardProperties (издатель)

Представляет параметры, доступные в публикации проекта или в мастере создания объекта макетов.
 


## <a name="example"></a>Пример

Используйте свойство **[Properties](wizard-properties-property-publisher.md)** с объектом **мастера** для возврата коллекции **WizardProperties** . Следующий пример отчетов по публикации проекта, связанного с активной публикации, отображение его имя и текущие настройки.
 

 

```
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 MsgBox "Publication Design associated with " _ 
 &amp; "current publication: " .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Примечание**  В зависимости от языковой версии Microsoft Publisher, используемая может появиться ошибка при использовании выше кода. В этом случае необходимо создать в обработчики ошибок для обхода ошибок. Для получения дополнительных сведений см **[Объект мастера](wizard-object-publisher.md)**.
 


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[FindPropertyById](wizardproperties-findpropertybyid-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](wizardproperties-application-property-publisher.md)|
|[Count](wizardproperties-count-property-publisher.md)|
|[Элемент](wizardproperties-item-property-publisher.md)|
|[Родительский раздел](wizardproperties-parent-property-publisher.md)|

