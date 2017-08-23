---
title: "Свойство WebNavigationBarSet.HorizontalButtonCount (издатель)"
keywords: vbapb10.chm8519687
f1_keywords: vbapb10.chm8519687
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.HorizontalButtonCount
ms.assetid: 2f6c5258-16c9-19fd-16c6-ea59c561e9de
ms.date: 06/08/2017
ms.openlocfilehash: 9488deefdb3c0f6784275cc1f055a1596ecc2a56
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsethorizontalbuttoncount-property-publisher"></a>Свойство WebNavigationBarSet.HorizontalButtonCount (издатель)

Задает или возвращает значение типа **Long** представляет число кнопок в каждой строке кнопки для набора панели навигации веб. Чтение и запись. **Длинные**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalButtonCount**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Возвращает значение «Доступ запрещен», если **IsHorizontal** = **значение False** для указанного объекта **WebNavigationBarSet** . Чтобы установить ориентацию панель навигации, равной **горизонтальной** сначала перед установкой **HorizontalButtonCount** , используйте метод **ChangeOrientation** .


## <a name="example"></a>Пример

В следующем примере возвращается первый панель навигации с активного документа изменения ориентации на **горизонтальную** при необходимости задает для свойства **HorizontalButtonCount** значение **3**и свойству **HorizontalAlignment** **pbnbAlignLeft**.


```vb
With ActiveDocument.WebNavigationBarSets(1) 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignRight 
End With
```


