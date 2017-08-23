---
title: "Свойство WebNavigationBarSet.Design (издатель)"
keywords: vbapb10.chm8519684
f1_keywords: vbapb10.chm8519684
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.Design
ms.assetid: 643d0b88-3b6d-65fd-7607-2f81c593a568
ms.date: 06/08/2017
ms.openlocfilehash: b9c3a2cf0f35dc0463fa22d7398318151db457a9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetdesign-property-publisher"></a>Свойство WebNavigationBarSet.Design (издатель)

Задает или возвращает константу **PbWizardNavBarDesign** , представляющее разработки указанный набор панели навигации веб. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Разработка**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="return-value"></a>Возвращаемое значение

PbWizardNavBarDesign


## <a name="remarks"></a>Заметки

Значение свойства **проекта** может иметь одно из **[PbWizardNavBarDesign](pbwizardnavbardesign-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере добавляется новой панели навигации веб задать для каждой страницы в активный документ, задает стиль кнопки для крупных и свойства проекта **pbnbDesignCapsule**.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newNavBar") 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleLarge 
 .Design = pbnbDesignCapsule 
End With
```


