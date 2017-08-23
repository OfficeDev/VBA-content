---
title: "Метод WebNavigationBarSet.ChangeOrientation (издатель)"
keywords: vbapb10.chm8519699
f1_keywords: vbapb10.chm8519699
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.ChangeOrientation
ms.assetid: bce05e9c-5b4a-f5a2-33a9-b40d4e05664f
ms.date: 06/08/2017
ms.openlocfilehash: 2d22eccd23ac8eeb255e7ba8f68be6290d6c166f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetchangeorientation-method-publisher"></a>Метод WebNavigationBarSet.ChangeOrientation (издатель)

Задает значение константы **PbNavBarOrientation** , представляющий выравнивания на панели навигации; вертикальной или горизонтальной.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ChangeOrientation** ( **_Ориентация_**)

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Ориентация|Обязательное свойство.| **PbNavBarOrientation**|

## <a name="remarks"></a>Заметки

Параметр ориентации может иметь одно из следующих **PbNavBarOrientation** константы, описанные в библиотеке типов, Microsoft Publisher.



| **pbNavBarOrientHorizontal**|| **pbNavBarOrientVertical**|

## <a name="example"></a>Пример

В следующем примере указывается, задайте объектную переменную на первом панель навигации в активный документ добавляется его для каждой страницы, изменяет ориентацию по горизонтали, задает горизонтальное выравнивание центра и наборы горизонтальной кнопки count — 4.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets(1) 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ChangeOrientation pbNavBarOrientHorizontal 
 .HorizontalAlignment = pbnbAlignCenter 
 .HorizontalButtonCount = 4 
End With
```


