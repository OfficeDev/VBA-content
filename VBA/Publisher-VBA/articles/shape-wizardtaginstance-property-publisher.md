---
title: "Свойство Shape.WizardTagInstance (издатель)"
keywords: vbapb10.chm2228339
f1_keywords: vbapb10.chm2228339
ms.prod: publisher
api_name: Publisher.Shape.WizardTagInstance
ms.assetid: 908d3f31-f277-7213-737e-9a946687bda7
ms.date: 06/08/2017
ms.openlocfilehash: 929882fbcc072311efe02bd80a07a0c6e2729922
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapewizardtaginstance-property-publisher"></a>Свойство Shape.WizardTagInstance (издатель)

Возвращает или задает **Long** , указывающее, экземпляр указанного фигуры, по сравнению с другими фигурами того же тег мастера. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WizardTagInstance**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Комбинация свойство **WizardTagInstance** и свойство **[WizardTag](shaperange-wizardtag-property-publisher.md)** однозначно определяет все фигуры в публикации.


## <a name="example"></a>Пример

Следующий пример отображает сведения об экземпляре тег мастера для всех фигур и мастер тега на странице один из активных публикации.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop 
 Debug.Print "Shape: " &; .Name 
 Debug.Print " Wizard tag: " &; .WizardTag 
 Debug.Print " Wizard tag instance: " _ 
 &; .WizardTagInstance 
 End With 
Next shpLoop
```


