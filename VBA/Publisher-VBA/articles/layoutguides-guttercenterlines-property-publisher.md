---
title: "Свойство LayoutGuides.GutterCenterlines (издатель)"
keywords: vbapb10.chm1114130
f1_keywords: vbapb10.chm1114130
ms.prod: publisher
api_name: Publisher.LayoutGuides.GutterCenterlines
ms.assetid: 7a5b1aef-85c7-548f-15e9-2c3b7327b439
ms.date: 06/08/2017
ms.openlocfilehash: 9b6b410f7ebed650a7d3ad98d3f4ddb1107bbbea
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesguttercenterlines-property-publisher"></a>Свойство LayoutGuides.GutterCenterlines (издатель)

Возвращает или задает значение, которое указывает, следует ли добавить строку центр между столбцов и строк из руководства по переплета в главную страницу. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GutterCenterlines**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Свойство **GutterCenterlines** можно использовать, только если ** [значение LayoutGuides.Rows](layoutguides-rows-property-publisher.md)** свойство или ** [LayoutGuides.Columns](layoutguides-columns-property-publisher.md)** больше, чем 1.

Если **значение True**, красной строки появляется в центре руководства по переплета. Если **значение False**, строка не отображается в центре руководства по переплета. Значение по умолчанию — **False**.


## <a name="example"></a>Пример

В следующем примере изменяется первая главная страница активная публикация трех строк, три столбца и представляют собой полностью руководства по переплета красной центр строки. Все страницы, добавлены к публикации после этого момента красной центр строки представляют собой полностью руководства по переплета.


```vb
Dim theMasterPage As page 
Dim theLayoutGuides As LayoutGuides 
 
Set theMasterPage = ActiveDocument.MasterPages(1) 
Set theLayoutGuides = theMasterPage.LayoutGuides 
 
With theLayoutGuides 
 .Rows = 3 
 .Columns = 3 
 .GutterCenterlines = True 
End With
```


