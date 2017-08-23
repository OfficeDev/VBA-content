---
title: "Метод WebNavigationBarSet.AddToEveryPage (издатель)"
keywords: vbapb10.chm8519698
f1_keywords: vbapb10.chm8519698
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.AddToEveryPage
ms.assetid: d36a3281-a313-084c-0ae9-7a981a7d9713
ms.date: 06/08/2017
ms.openlocfilehash: 2220313d87581ba74a0a91be16b36209c1bc5626
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetaddtoeverypage-method-publisher"></a>Метод WebNavigationBarSet.AddToEveryPage (издатель)

Добавляет **ShapeRange** из типа **pbWebNavigationBar** для каждой страницы текущего документа.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddToEveryPage** ( **_Слева_**, **_Top_**, **_Width_**)

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Слева|Обязательное свойство.| **Variant**|Задайте положение левого края фигуры, представляющий панель навигации.|
|Вверх|Обязательное свойство.| **Variant**|Задайте положение верхнего края фигуры, представляющий панель навигации.|
|Width|Необязательный| **Variant**|Задать ширину фигуры, представляющий панель навигации.|

### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="remarks"></a>Заметки

Указанный набор панель навигации Web должен существовать до вызова этого метода. 


## <a name="example"></a>Пример

В следующем примере добавляется набора именованные «WebNavBarSet1» панель навигации в верхней части каждой страницы в активный документ.


```vb
ActiveDocument.WebNavigationBarSets("WebNavBarSet1") _ 
 .AddToEveryPage Left:=10, Top:=20 

```

В следующем примере добавляется новый панель навигации задайте в активный документ и добавляет его на все страницы публикации.




```vb
Dim objWebNavBarSet As WebNavigationBarSet 
 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.AddSet( _ 
 Name:="WebNavBarSet1", _ 
 Design:=pbnbDesignTopLine, _ 
 AutoUpdate:=True) 
 
objWebNavBarSet.AddToEveryPage Left:=50, Top:=10, Width:=500
```


