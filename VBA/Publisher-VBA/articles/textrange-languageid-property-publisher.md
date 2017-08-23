---
title: "Свойство TextRange.LanguageID (издатель)"
keywords: vbapb10.chm5308471
f1_keywords: vbapb10.chm5308471
ms.prod: publisher
api_name: Publisher.TextRange.LanguageID
ms.assetid: 1007c821-cafd-0cb3-94f4-4ac25decad30
ms.date: 06/08/2017
ms.openlocfilehash: b7a5841aa468dd4af818708e93b508e7c0cb4cab
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangelanguageid-property-publisher"></a>Свойство TextRange.LanguageID (издатель)

Возвращает или задает значение константы **MsoLanguageID** , представляющий язык для указанного объекта. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LanguageID**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

MsoLanguageID


## <a name="remarks"></a>Заметки

Значение свойства **LanguageID** может иметь одно из ** [MsoLanguageID](http://msdn.microsoft.com/library/65ea40f0-9a09-3d76-1519-4acddcc5f367%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере форматов указанного выбора на французском языке. В этом примере предполагается, что курсор находится в текстовом поле.


```vb
Sub SetLanguage() 
 Selection.TextRange.LanguageID = msoLanguageIDFrench 
End Sub
```


