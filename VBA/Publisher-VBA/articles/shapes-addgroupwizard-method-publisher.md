---
title: "Метод Shapes.AddGroupWizard (издатель)"
keywords: vbapb10.chm2162727
f1_keywords: vbapb10.chm2162727
ms.prod: publisher
api_name: Publisher.Shapes.AddGroupWizard
ms.assetid: 5a84f055-7f30-0757-f507-40ee34b214f4
ms.date: 06/08/2017
ms.openlocfilehash: 27d22ac1bd7211ef75ddd4f903700e1d7753d799
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddgroupwizard-method-publisher"></a>Метод Shapes.AddGroupWizard (издатель)

Добавление объекта **Shape** , представляющий объект макетов публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddGroupWizard** ( **_Мастер_**, **_слева_**, **_сверху_**, **_Ширина_**, **_Высота_**, **_разработки_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Мастер|Обязательное свойство.| **PbWizardGroup**|Тип объекта макетов, чтобы добавить к публикации.|
|Слева|Обязательное свойство.| **Variant**|Позиция левого края макетов объектов относительно левого края страницы, заданная в пунктах.|
|Вверх|Обязательное свойство.| **Variant**|Позиция верхнего края макетов объектов относительно верхнего края страницы, заданная в пунктах.|
|Width|Необязательный| **Variant**|Ширина новый объект макетов.|
|Height|Необязательный| **Variant**|Высота новый объект макетов.|
|Разработка|Необязательный| **Длинный**|Разработка объекта будет добавлена.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Параметр мастера может иметь одно из **[PbWizardGroup](pbwizardgroup-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере добавляется веб-оглавление active публикацию.


```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddGroupWizard Wizard:=pbWizardGroupTableOfContents, _ 
 Left:=100, Top:=100
```


