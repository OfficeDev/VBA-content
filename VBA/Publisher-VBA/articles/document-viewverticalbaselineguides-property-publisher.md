---
title: "Свойство Document.ViewVerticalBaseLineGuides (издатель)"
keywords: vbapb10.chm196729
f1_keywords: vbapb10.chm196729
ms.prod: publisher
api_name: Publisher.Document.ViewVerticalBaseLineGuides
ms.assetid: 711335ab-237b-65a2-534a-7635cfba474e
ms.date: 06/08/2017
ms.openlocfilehash: 8d387d5d07cda91ff59e981886e2794a2651b5e3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentviewverticalbaselineguides-property-publisher"></a>Свойство Document.ViewVerticalBaseLineGuides (издатель)

Задает или возвращает значение **типа Boolean** , представляет ли вертикальная направляющие видны в указанном объекте **документа** . **Значение true,** если они будут отображаться. **Значение false,** Если эти атрибуты не видны. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ViewVerticalBaseLineGuides**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

По умолчанию для этого свойства имеет **значение False**.


## <a name="example"></a>Пример

В следующем примере создается руководства по вертикали базового видимы в активный документ.


```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
objDocument.ViewVerticalBaseLineGuides = True 

```


