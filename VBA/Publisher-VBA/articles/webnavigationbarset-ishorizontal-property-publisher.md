---
title: "Свойство WebNavigationBarSet.IsHorizontal (издатель)"
keywords: vbapb10.chm8519686
f1_keywords: vbapb10.chm8519686
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.IsHorizontal
ms.assetid: d3bbb0b0-8d06-7d46-1ef7-fef0a3e846b7
ms.date: 06/08/2017
ms.openlocfilehash: fcf7d3f4ac8b1d62f61aeab460e4ff7caeb3d8ad
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetishorizontal-property-publisher"></a>Свойство WebNavigationBarSet.IsHorizontal (издатель)

 **Значение true,** Если для указанного **WebNavigationBarSet** горизонтальное расположение. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsHorizontal**

 переменная _expression_A, представляющий объект **WebNavigationBarSet** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство используется для определения ориентации установлено до установки некоторых свойств, которые предполагается горизонтальное расположение, такими как свойство **HorizontalAlignment** панели навигации.


## <a name="example"></a>Пример

В этом примере добавляет первой панели навигации в коллекции **WebNavigationBarSets** активного документа для каждой страницы и затем задает стиль кнопки для **малых**. Для определения, является ли набор панель навигации горизонтальной выполняется проверка. Если он не установлен, вызывается метод **ChangeOrientation** , ориентации задано значение **Горизонтальная**. После ориентирована на панели навигации по горизонтали, count горизонтальной кнопки задано значение **3** , горизонтальное выравнивание кнопок задано значение **слева**.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets(1) 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleSmall 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignLeft 
End With
```


