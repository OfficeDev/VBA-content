---
title: "Метод Shapes.AddCatalogMergeArea (издатель)"
keywords: vbapb10.chm2162752
f1_keywords: vbapb10.chm2162752
ms.prod: publisher
api_name: Publisher.Shapes.AddCatalogMergeArea
ms.assetid: 4af86b99-5a3a-b9f3-d269-16d635d35c83
ms.date: 06/08/2017
ms.openlocfilehash: 74e15744dcac7520600bdc46ec7cdd277ff745db
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddcatalogmergearea-method-publisher"></a>Метод Shapes.AddCatalogMergeArea (издатель)

Добавляет объект **фигуры** , представляющий область указанной публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddCatalogMergeArea**

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Только одна область можно добавить на страницу публикации. Как правило публикации будут иметь только одну область.

Несмотря на то, что можно добавить одну область на каждой странице публикации, можно только подключения к источнику данных для публикации. Какие данные объединяются определяет, какие области данных на активной странице и полей данных, которые он содержит.


 **Примечание**  Перед подключением к источнику данных, необходимо добавить область страницы публикации.

Используйте метод **[AddToCatalogMergeArea](shape-addtocatalogmergearea-method-publisher.md)** объектов **[фигуры](shape-object-publisher.md)** или **[ShapeRange](shaperange-object-publisher.md)** добавление фигур в области объединения в каталог.

Используйте метод **[вставки](mailmergedatafield-insert-method-publisher.md)** коллекции **[MailMergeDataFields](mailmergedatafields-object-publisher.md)** Добавление поля данных изображения в области публикации.

Используйте метод **[InsertMailMergeField](textrange-insertmailmergefield-method-publisher.md)** объекта **[TextRange](textrange-object-publisher.md)** Добавление текстового поля данных в текстовом поле в области публикации.

Используйте метод **[RemoveCatalogMergeArea](shape-removecatalogmergearea-method-publisher.md)** объекта **[Shape](shape-object-publisher.md)** , чтобы удалить область из публикации.

Этот метод соответствует выбору объединения в каталог в **Шаг 1: Выбор типа слияния** **почты и каталог**.


## <a name="example"></a>Пример

В следующем примере добавляется область объединения в каталог на первой странице указанной публикации.


```vb
ThisDocument.Pages(1).Shapes.AddCatalogMergeArea
```


