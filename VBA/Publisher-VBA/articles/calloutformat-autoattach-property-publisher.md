---
title: "Свойство CalloutFormat.AutoAttach (издатель)"
keywords: vbapb10.chm2490626
f1_keywords: vbapb10.chm2490626
ms.prod: publisher
api_name: Publisher.CalloutFormat.AutoAttach
ms.assetid: 893303d8-97fe-9eea-8d6e-d9110c75ee84
ms.date: 06/08/2017
ms.openlocfilehash: d9bb5a593152f67b04b82b505a2381d59a80d581
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatautoattach-property-publisher"></a>Свойство CalloutFormat.AutoAttach (издатель)

Возвращает или задает константой **MsoTriState**, указывающее, является ли место, где линии выноски подключает текстовое поле выноски изменяется в зависимости от того, является ли origin линии выноски (где выноски указывает) слева или справа от надписи выноски. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutoAttach**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **AutoAttach** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Если значение этого свойства **msoTrue**, текстовое поле справа от источника, и измеряется в нижней части текстового поля при текстовом поле слева от происхождения при значение перетаскивания (вертикальной расстояние от края текстовое поле выноски в то место, где подключает линии выноски) отсчитывается от верхней части текстового поля. Если значение этого свойства **msoFalse**, значение перетаскивания всегда отсчитывается от верхней части текстового поля, независимо от того, относительное расположение текстовое поле и источник. Метод [CustomDrop](calloutformat-customdrop-method-publisher.md)используется для установки значения раскрывающегося и используйте свойство [поместите](calloutformat-drop-property-publisher.md)для возврата значения размещения сообщений.

Установка для этого свойства влияет на выноске только в том случае, если он имеет явным образом установлен перетащите значение, то есть, если значение свойства [DropType](calloutformat-droptype-property-publisher.md) **msoCalloutDropCustom**. По умолчанию выноски задать размещения значения при создании.


## <a name="example"></a>Пример

В этом примере добавляется два выноски для первой страницы. Один из выноски подключен автоматически, а другое — не. При изменении происхождение строки выноски для автоматически вложенные выноски справа от вложенные текстовое поле, изменяется положение текстовое поле. Выноска, не подключенного автоматически не отображает это поведение.


```vb
With ActivePublication.Pages(1).Shapes 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=170, Width:=200, Height:=50) 
 .TextFrame.TextRange.Text = "auto-attached" 
 .Callout.AutoAttach = msoTrue 
 End With 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=350, Width:=200, Height:=50) 
 .TextFrame.TextRange.Text = "not auto-attached" 
 .Callout.AutoAttach = msoFalse 
 End With 
End With 

```


