---
title: "Свойство Page.IsWizardPage (издатель)"
keywords: vbapb10.chm393271
f1_keywords: vbapb10.chm393271
ms.prod: publisher
api_name: Publisher.Page.IsWizardPage
ms.assetid: 09c1352d-6760-ad54-aa95-211727c968b3
ms.date: 06/08/2017
ms.openlocfilehash: 96d95e02c9b0e93d6f91f99f91a28b54a0ef1cd5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageiswizardpage-property-publisher"></a>Свойство Page.IsWizardPage (издатель)

Возвращает **значение True** , если это, страница мастера Microsoft Publisher. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsWizardPage**

 переменная _expression_A, представляющий объект **страницы** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Страницы мастера являются типы специальных страниц для определенных типов мастеров Publisher (например, бюллетени, каталоги и веб-мастера), которые можно вставить в публикации.

Используйте свойство **[Мастер](page-wizard-property-publisher.md)** объекта **[Page](page-object-publisher.md)** для доступа к мастеру для указанной страницы.


## <a name="example"></a>Пример

Следующий пример проверяет, чтобы определить, является ли указанный страницы страница мастера. Если он установлен, будут возвращены определенных свойств мастера.


```vb
 With ActiveDocument.Pages(1) 
 If .IsWizardPage = True Then 
 
 With .Wizard 
 Debug.Print .Name 
 Debug.Print .Properties(1).Name 
 Debug.Print .Properties(1).CurrentValueId 
 End With 
 
 End If 
 End With
```


