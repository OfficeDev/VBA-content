---
title: "Метод Shape.AddToCatalogMergeArea (издатель)"
keywords: vbapb10.chm5308688
f1_keywords: vbapb10.chm5308688
ms.prod: publisher
api_name: Publisher.Shape.AddToCatalogMergeArea
ms.assetid: 4178d286-045f-a7b6-86b6-710bed10e824
ms.date: 06/08/2017
ms.openlocfilehash: 66ef7d3a1ea86c2d3ab917fbe3f60c60faf1c44a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# Метод Shape.AddToCatalogMergeArea (издатель)

Добавляет указанный фигуры или фигур области страницы публикации.


## Синтаксис

 _выражение_. **AddToCatalogMergeArea**

 переменная _expression_A, представляющий объект **фигуры** .


### Возвращаемое значение

Значение Nothing


## Заметки

Область данных автоматически изменяется, чтобы вместить объекты, размер которых превышает области объединения или находятся вне области данных, после их добавления.

Метод **AddToCatalogMergeArea** не применяется для объединения поля данных:


- Используйте метод **[вставки](mailmergedatafield-insert-method-publisher.md)** коллекции **[MailMergeDataFields](mailmergedatafields-object-publisher.md)** Добавление поля данных изображения в области страницы публикации.
    
- Используйте метод **[InsertMailMergeField](textrange-insertmailmergefield-method-publisher.md)** объекта **[TextRange](textrange-object-publisher.md)** Добавление текстового поля данных в текстовом поле.
    


Обратите внимание на то, чтобы добавить текстовое поле, которое будет содержать текстовых полей данных в области объединения в каталог, используйте метод **AddToCatalogMergeArea** .


## Пример

Следующий пример добавляет прямоугольник области данных на первой странице указанной публикации. В этом примере предполагается, что область объединения в каталог был добавлен к первой странице.


```vb
ThisDocument.Pages(1).Shapes.AddShape(1, 80, 75, 450, 125).AddToCatalogMergeArea
```


