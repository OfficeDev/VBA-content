---
title: "Метод Document.FindShapeByWizardTag (издатель)"
keywords: vbapb10.chm196690
f1_keywords: vbapb10.chm196690
ms.prod: publisher
api_name: Publisher.Document.FindShapeByWizardTag
ms.assetid: c6db9ba7-15b0-e8f0-1ed2-08b6e978c948
ms.date: 06/08/2017
ms.openlocfilehash: 5296cf64352ff077e1f95b73d6aef482a4f03b7a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentfindshapebywizardtag-method-publisher"></a>Метод Document.FindShapeByWizardTag (издатель)

Возвращает объект **ShapeRange** , представляющая одного или всех фигур в публикации с помощью мастера и с тегом указанного мастера.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FindShapeByWizardTag** ( **_WizardTag_**, **_экземпляр_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|WizardTag|Обязательное свойство.| **PbWizardTag**|Задает тег мастера для поиска.|
|Экземпляр|Необязательный| **Длинный**|Указывает, какой экземпляр фигуры с тегом указанного мастера возвращается. Для экземпляра равно n, n-й экземпляра фигуры с тегом указанного мастера возвращается. Если значение не для экземпляра не указан, возвращаются все фигур с помощью мастера указанный тег.|

### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="remarks"></a>Заметки

Параметр WizardTag может иметь одно из **[PbWizardTag](pbwizardtag-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В следующем примере выполняется поиск второго экземпляра фигуры с тегом мастер **pbWizardDate** и присваивается переменной.


```vb
Dim shpWizardTag As Shape 
 
Set shpWizardTag = ActiveDocument._ 
 FindShapeByWizardTag(WizardTag:=pbWizardDate, Instance:=2)
```


