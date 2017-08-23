---
title: "Метод Tags.Add (издатель)"
keywords: vbapb10.chm4653060
f1_keywords: vbapb10.chm4653060
ms.prod: publisher
api_name: Publisher.Tags.Add
ms.assetid: 78602ccc-8198-1183-4775-fe626eb8b5af
ms.date: 06/08/2017
ms.openlocfilehash: 66d39411df5eb7d789d671f13a9580bb552daa65
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tagsadd-method-publisher"></a>Метод Tags.Add (издатель)

Добавляет новый объект **тега** на указанный объект **теги** и возвращает новый объект **тега** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Имя_**, **_значение_**)

 переменная _expression_A, представляет собой объект- **теги** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|Обязательное свойство.| **String**|Имя тега для добавления. Если тег с таким же именем уже существует, возникает ошибка.|
|Значение|Обязательное свойство.| **Variant**|Значение, задаваемое в тег.|

### <a name="return-value"></a>Возвращаемое значение

Тег


## <a name="example"></a>Пример

В следующем примере добавляется тег фигуру один на один из активных публикации страницы.


```vb
Dim tagNew As Tag 
 
Set tagNew = ActiveDocument.Pages(1).Shapes(1).Tags _ 
 .Add(Name:="required", Value:="yes")
```


