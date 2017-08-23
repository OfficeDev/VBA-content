---
title: "Свойство Document.ViewHorizontalBaseLineGuides (издатель)"
keywords: vbapb10.chm196728
f1_keywords: vbapb10.chm196728
ms.prod: publisher
api_name: Publisher.Document.ViewHorizontalBaseLineGuides
ms.assetid: e5471313-38e0-9454-04af-4c85d976b312
ms.date: 06/08/2017
ms.openlocfilehash: fbe93d93ce9c183118ceea6fb6ff26994e657b24
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentviewhorizontalbaselineguides-property-publisher"></a>Свойство Document.ViewHorizontalBaseLineGuides (издатель)

Задает или возвращает значение **типа Boolean** , представляет ли горизонтальные направляющие отображаются в указанном объекте **документа** . **Значение true,** если они будут отображаться. **Значение false,** Если эти атрибуты не видны. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. ****

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

По умолчанию для этого свойства имеет **значение False**.


## <a name="example"></a>Пример

В следующем примере создается руководства по горизонтали базового видимы в активный документ.


```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
objDocument.ViewHorizontalBaseLineGuides = True 

```


