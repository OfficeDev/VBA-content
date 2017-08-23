---
title: "Свойство Page.Wizard (издатель)"
keywords: vbapb10.chm393234
f1_keywords: vbapb10.chm393234
ms.prod: publisher
api_name: Publisher.Page.Wizard
ms.assetid: 05cf1482-bde5-9ea2-4099-69a56a2dc61a
ms.date: 06/08/2017
ms.openlocfilehash: 4d0c2937fd84a75420cd21b9e715634bde1ed933
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagewizard-property-publisher"></a>Свойство Page.Wizard (издатель)

Возвращает объект **[мастера](wizard-object-publisher.md)** , представляющий макет публикации, связанные с указанной публикации или мастера, связанного с указанным объектом макетов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Мастер**

 переменная _expression_A, представляющий объект **Page** .


## <a name="remarks"></a>Заметки

При обращении к свойству **мастера** из объекта **документа** или **страницы** при указанной публикации не связан с любой макет публикации, возникает ошибка. При доступе к свойству **мастера** из объекта **фигуры** или **ShapeRange** Если указанный объект не является объектом макетов, возникает ошибка.


## <a name="example"></a>Пример

Следующий пример отчетов по публикации проекта, связанного с активной публикации, отображение его имя и текущие настройки.


```vb
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 Debug.Print "Publication design associated with " _ 
 &; "current publication: " _ 
 &; .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Setting: " _ 
 &; .Name &; " = " &; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Примечание**  В зависимости от языковой версии Publisher, используемая может появиться ошибка при использовании выше кода. В этом случае необходимо создать в обработчики ошибок для обхода ошибок. Для получения дополнительных сведений см **[Мастер](wizard-object-publisher.md)** объекта.


