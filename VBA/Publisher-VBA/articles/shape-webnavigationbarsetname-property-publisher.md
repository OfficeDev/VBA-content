---
title: "Свойство Shape.WebNavigationBarSetName (издатель)"
keywords: vbapb10.chm5308677
f1_keywords: vbapb10.chm5308677
ms.prod: publisher
api_name: Publisher.Shape.WebNavigationBarSetName
ms.assetid: 0d9abe17-6936-562b-9210-5f092d13f215
ms.date: 06/08/2017
ms.openlocfilehash: 4f92c58275b2f05749f8e4008e9d871decdb70a0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapewebnavigationbarsetname-property-publisher"></a>Свойство Shape.WebNavigationBarSetName (издатель)

Возвращает **строку** , представляющую имя указанной фигуры является экземпляром набор панели навигации веб. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebNavigationBarSetName**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Это свойство доступно только для фигур, которые представляют экземпляр набора панель навигации Web. Свойство **[типа](shape-type-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** для определения, если фигуры представляет экземпляр объекта набора панель навигации Web.

Свойство **WebNavigationBarSetName** возвращает имя объекта **[WebNavigationBarSet](webnavigationbarset-object-publisher.md)** . Несколько страниц веб-публикации может иметь форму, представляющую экземпляр же панель навигации задать. Изменения, внесенные в объект **WebNavigationBarSet** отражаются в все фигуры, представляющие экземпляры этот набор панель навигации Web.


## <a name="example"></a>Пример

Следующий пример проверяет позволяет определить, какие фигуры на первой странице активного документа представляют экземпляры Web панелей навигации. Для каждой такой фигуры найден панель навигации, который представляет экземпляр объекта задано значение автоматическое обновление.


```vb
Sub SetWebBarsToAutoUpdate() 
 
Dim shpLoop As Shape 
Dim strWebNavBarName As String 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = pbWebNavigationBar Then 
 
 strWebNavBarName = shpLoop.WebNavigationBarSetName 
 With ActiveDocument.WebNavigationBarSets(strWebNavBarName) 
 .AutoUpdate = True 
 End With 
 
 End If 
Next 
 
End Sub
```


