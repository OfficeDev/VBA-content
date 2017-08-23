---
title: "Свойство WebCheckBox.Value (издатель)"
keywords: vbapb10.chm4325381
f1_keywords: vbapb10.chm4325381
ms.prod: publisher
api_name: Publisher.WebCheckBox.Value
ms.assetid: 9fd50cd5-ecf3-30b7-c8a9-6b64b106eaec
ms.date: 06/08/2017
ms.openlocfilehash: a8b06b34fa4647725c04cec9757c98a242eebb6d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcheckboxvalue-property-publisher"></a>Свойство WebCheckBox.Value (издатель)

Возвращает или задает **строку** , представляющую значение Web флажок или переключатель. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Значение**

 переменная _expression_A, представляет собой объект- **WebCheckBox** .


## <a name="example"></a>Пример

В этом примере создается новый флажок веб-элемент управления и присваивает ему имя и значение указывает, что установлен флажок исходное состояние.


```vb
Sub CreateWebButton() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCheckBox, Left:=72, _ 
 Top:=72, Width:=100, Height:=50) 
 .Name = "ControlBox" 
 With .WebCheckBox 
 .Value = "This is a check box." 
 .Selected = msoTrue 
 End With 
 End With 
End Sub
```


