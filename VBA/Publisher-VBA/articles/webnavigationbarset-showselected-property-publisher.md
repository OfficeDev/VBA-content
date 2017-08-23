---
title: "Свойство WebNavigationBarSet.ShowSelected (издатель)"
keywords: vbapb10.chm8519696
f1_keywords: vbapb10.chm8519696
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.ShowSelected
ms.assetid: c8229f03-a043-a280-84f9-f75a430c3903
ms.date: 06/08/2017
ms.openlocfilehash: 6081ba8f46bb9e2fb0002dcf37fa275e91c1e236
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetshowselected-property-publisher"></a>Свойство WebNavigationBarSet.ShowSelected (издатель)

 **Значение true,** Если для указанного объекта **WebNavigationBarSet** будет выделена выбранной кнопки. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShowSelected**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В следующем примере добавляется новый панель навигации в активный документ добавляется на все страницы и затем устанавливает для свойства **ShowSelected** значение **False** , чтобы выбранной кнопки не выделяются в панели навигации.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newNavBar") 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleSmall 
 .ShowSelected = False 
End With
```


