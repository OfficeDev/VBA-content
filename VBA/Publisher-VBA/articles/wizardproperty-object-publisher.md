---
title: "Объект WizardProperty (издатель)"
keywords: vbapb10.chm1638399
f1_keywords: vbapb10.chm1638399
ms.prod: publisher
api_name: Publisher.WizardProperty
ms.assetid: 9f059422-5454-1902-a092-76e21e36a3f7
ms.date: 06/08/2017
ms.openlocfilehash: 59eb119b1655ae5f031acf569224b10532a93ab4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardproperty-object-publisher"></a>Объект WizardProperty (издатель)

Представляет параметр, который является частью определенного публикации проекта или мастер объектов макетов.
 


## <a name="example"></a>Пример

Используйте свойство **[Item](wizardproperties-item-property-publisher.md)** или метод **[FindByPropertyID](wizardproperties-findpropertybyid-method-publisher.md)** с коллекцией **WizardProperties** для возврата объекта **WizardProperty** . Следующий пример отчетов по публикации проекта, связанного с активной публикации, отображение его имя и текущие настройки.
 

 

```
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
 Debug.Print " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Примечание**  В зависимости от языковой версии Microsoft Publisher, используемая может появиться ошибка при использовании выше кода. В этом случае необходимо создать в обработчики ошибок для обхода ошибок. Для получения дополнительных сведений см **[Объект мастера](wizard-object-publisher.md)**.
 


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](wizardproperty-application-property-publisher.md)|
|[CurrentValueId](wizardproperty-currentvalueid-property-publisher.md)|
|[Включено](wizardproperty-enabled-property-publisher.md)|
|[ID](wizardproperty-id-property-publisher.md)|
|[Name](wizardproperty-name-property-publisher.md)|
|[Родительский раздел](wizardproperty-parent-property-publisher.md)|
|[Значения](wizardproperty-values-property-publisher.md)|

