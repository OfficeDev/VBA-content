---
title: "Свойство CalloutFormat.Length (издатель)"
keywords: vbapb10.chm2490632
f1_keywords: vbapb10.chm2490632
ms.prod: publisher
api_name: Publisher.CalloutFormat.Length
ms.assetid: 878fdb7b-fca6-49b6-1ec0-143243ce014c
ms.date: 06/08/2017
ms.openlocfilehash: f347a2adc084919e0d3707ef65c76b2d2f346c47
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatlength-property-publisher"></a>Свойство CalloutFormat.Length (издатель)

Возвращает **Variant** , указывающее длину (в пунктах) первый сегмент линии выноски (сегмент, подключенного к поле выноски) Если свойство **[AutoLength](calloutformat-autolength-property-publisher.md)** указанного выноски установлено значение **False**. В противном случае возникает ошибка. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Длина**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


## <a name="remarks"></a>Заметки

Это свойство применяется только к выноски, чьи строки состоят из нескольких сегментов (типы **msoCalloutThree** и **msoCalloutFour**).

Используйте метод **[CustomLength](calloutformat-customlength-method-publisher.md)** для задания значения этого свойства.


## <a name="example"></a>Пример

Если первый отрезок на выноске с именем co1 имеет фиксированную длину, в этом примере указывается, что длина первого отрезка линии в выноске, с именем co2, также исправляются в заданной длины. Для обеспечения работы примера оба выноски должна иметь сегмент несколько строк.


```vb
Dim len1 As Single 
 
With ActiveDocument.Pages(1).Shapes 
 With .Item("co1").Callout 
 If Not .AutoLength Then len1 = .Length 
 End With 
 If len1 Then .Item("co2").Callout _ 
 .CustomLength Length:=len1 
End With
```


