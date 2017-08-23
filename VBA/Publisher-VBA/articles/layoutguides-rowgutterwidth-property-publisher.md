---
title: "Свойство LayoutGuides.RowGutterWidth (издатель)"
keywords: vbapb10.chm1114129
f1_keywords: vbapb10.chm1114129
ms.prod: publisher
api_name: Publisher.LayoutGuides.RowGutterWidth
ms.assetid: a7629683-68d2-4953-4c95-7e79e431f9c4
ms.date: 06/08/2017
ms.openlocfilehash: 9071ddcd0786fb881673574cf6a4a96e8330f2f4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesrowgutterwidth-property-publisher"></a>Свойство LayoutGuides.RowGutterWidth (издатель)

Возвращает или задает ширину переплета строк, используемые в объекте **LayoutGuides** для помощи в процессе с макетом элементы дизайна. Чтение и запись **одного**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RowGutterWidth**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Ширина по умолчанию переплета строки — 0,4 дюйма.


## <a name="example"></a>Пример

В следующем примере изменяется вторая главная страница active публикации, чтобы она имела четыре строки и четыре столбца, ширина переплета строки 0,75 дюйма, ширина столбца переплета 0,5 дюйма и центр строки в переплета. Новые страницы добавлены к публикации, используйте второй главную страницу как шаблон будет этих свойств.


```vb
Dim theMasterPage As page 
Dim theLayoutGuides As LayoutGuides 
 
Set theMasterPage = ActiveDocument.MasterPages(2) 
Set theLayoutGuides = theMasterPage.LayoutGuides 
 
With theLayoutGuides 
 .Rows = 4 
 .Columns = 4 
 .RowGutterWidth = Application.InchesToPoints(0.75) 
 .ColumnGutterWidth = Application.InchesToPoints(0.5) 
 .GutterCenterlines = True 
End With
```


