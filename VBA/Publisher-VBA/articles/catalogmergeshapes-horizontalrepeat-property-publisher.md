---
title: "Свойство CatalogMergeShapes.HorizontalRepeat (издатель)"
keywords: vbapb10.chm8388613
f1_keywords: vbapb10.chm8388613
ms.prod: publisher
api_name: Publisher.CatalogMergeShapes.HorizontalRepeat
ms.assetid: 1c3f1093-294f-e7b3-02ca-803ce7437d49
ms.date: 06/08/2017
ms.openlocfilehash: dc0a75900163e49be51d0e13a02c4e9ee7015c17
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="catalogmergeshapeshorizontalrepeat-property-publisher"></a>Свойство CatalogMergeShapes.HorizontalRepeat (издатель)

Возвращает значение типа **Long** , представляющее количество раз, когда область данных будет повторите через конечной страницы публикации, при выполнении объединение в каталог. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalRepeat**

 переменная _expression_A, представляет собой объект- **CatalogMergeShapes** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Когда выполняется объединение в каталог, область данных повторяется один раз для каждого выбранного записи в указанный источник данных.

Сколько раз повторяет области данных на странице, определяется ширины области. Свойство **[Width](shape-width-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** для возвращения или задания горизонтальной размера области данных.

Свойство **[VerticalRepeat](catalogmergeshapes-verticalrepeat-property-publisher.md)** объекта **[CatalogMergeShapes](catalogmergeshapes-object-publisher.md)** представляет число раз, когда область данных по вертикали повторяет вниз конечной страницы публикации.


## <a name="example"></a>Пример

В следующем примере возвращается количество раз, когда область данных будет повторяющиеся по горизонтали и по вертикали на конечной страницы публикации, когда выполняется объединение в каталог. В этом примере предполагается, что область данных является первой фигуры на первой странице указанной публикации.


```vb
Sub CatalogMergeDimensions() 
 
 With ThisDocument.Pages(1).Shapes(1) 
 Debug.Print .Width 
 Debug.Print .CatalogMergeItems.HorizontalRepeat 
 Debug.Print .Height 
 Debug.Print .CatalogMergeItems.VerticalRepeat 
 End With 
 
End Sub
```


