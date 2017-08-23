---
title: "Метод MailMergeDataSource.SetAllIncludedFlags (издатель)"
keywords: vbapb10.chm6291481
f1_keywords: vbapb10.chm6291481
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.SetAllIncludedFlags
ms.assetid: ab668e95-55ac-fcbd-19c9-3c13fe3aa995
ms.date: 06/08/2017
ms.openlocfilehash: 55daaa431fef7329bddced245c1753a2540aafd2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcesetallincludedflags-method-publisher"></a>Метод MailMergeDataSource.SetAllIncludedFlags (издатель)

 **Значение true,** Чтобы включить все записи источника данных при слиянии.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetAllIncludedFlags** ( **_Включено_**)

 переменная _expression_A, представляющий объект **вывода** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Включенные|Обязательное свойство.| **Boolean**| **Значение true,** Чтобы включить все записи источника данных при слиянии. **Значение false,** чтобы исключить все записи источника данных из слияния почты.|

## <a name="remarks"></a>Заметки

Позволяет указать отдельные записи в источник данных, чтобы быть включены или исключены из слияния почты, с помощью свойства **[включено](mailmergedatasource-included-property-publisher.md)** .


## <a name="example"></a>Пример

В этом примере все записи в источнике данных, содержащее поле недопустимый адреса, задает комментарий, почему он является недопустимым и исключает все записи из слияния почты.


```vb
Sub FlagAllRecords() 
 With ActiveDocument.MailMerge.DataSource 
 .SetAllErrorFlags Invalid:=True, InvalidComment:= _ 
 "All records in the data source have only 5-" _ 
 &; "digit ZIP Codes. Need 5+4 digit ZIP Codes." 
 .SetAllIncludedFlags Included:=False 
 End With 
End Sub
```


