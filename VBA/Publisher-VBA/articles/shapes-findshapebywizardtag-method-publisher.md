---
title: "Метод Shapes.FindShapeByWizardTag (издатель)"
keywords: vbapb10.chm2162728
f1_keywords: vbapb10.chm2162728
ms.prod: publisher
api_name: Publisher.Shapes.FindShapeByWizardTag
ms.assetid: f1018f3a-4f8f-2686-ac58-6eee8827c743
ms.date: 06/08/2017
ms.openlocfilehash: bf790182f9f836ad277c3e08188db27958691627
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesfindshapebywizardtag-method-publisher"></a>Метод Shapes.FindShapeByWizardTag (издатель)

Возвращает объект **ShapeRange** , представляющая одного или всех фигур в публикации с помощью мастера и с тегом указанного мастера.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FindShapeByWizardTag** ( **_WizardTag_**, **_экземпляр_**)

 переменная _expression_A, представляет собой объект- **фигур** .


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
 
Set shpWizardTag = ActiveDocument.FindShapeByWizardTag(WizardTag:=pbWizardDate, Instance:=2)
```


