---
title: "Свойство WizardProperty.ID (издатель)"
keywords: vbapb10.chm1572867
f1_keywords: vbapb10.chm1572867
ms.prod: publisher
api_name: Publisher.WizardProperty.ID
ms.assetid: 2827af5d-d002-029b-7f93-26befe459229
ms.date: 06/08/2017
ms.openlocfilehash: cb7a625a7d0444f5ef2c3b3d969d5e03bcb377eb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardpropertyid-property-publisher"></a>Свойство WizardProperty.ID (издатель)

Возвращает значение типа **Long** , представляющее тип фигуры, диапазона фигур или свойство, тип или значение мастера. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Идентификатор**

 переменная _expression_A, представляет собой объект- **WizardProperty** .


## <a name="example"></a>Пример

В этом примере тип для каждой фигуры на первой странице active публикации.


```vb
Sub ShapeID() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 MsgBox shp.ID 
 Next shp 
End Sub
```


