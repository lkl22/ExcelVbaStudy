# Dictionary 对象

存储数据键/item对的对象。

* [语法](#语法)
* [Remarks](#Remarks)
* [方法](#方法)
* [属性](#属性)

## 语法
```vba
Scripting.Dictionary
```

## <a name="Remarks">Remarks</a>

`Dictionary` 对象是 PERL 关联阵列的等效项。 可以是任何形式的数据的item存储在阵列中。 每个 item 都与唯一的键关联。 键用于检索单个item, 它通常是整数或字符串, 但除了数组之外, 还可以是任何其他项。

下面的代码演示如何创建 `Dictionary` 对象。
```vba
Dim d                   'Create a variable
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
...
```

## 方法

|方法	|说明
|---|---
|[Add](#Addmethod)	|将新的键/item对添加到 Dictionary 对象。
|[Exists](#Existsmethod)	|返回一个布尔值, 该值指示是否在 Dictionary 对象中存在指定的键。
|[Items](#Itemsmethod)	|返回一个由 Dictionary 对象中的所有item组成的数组。
|[Keys](#Keysmethod)	|返回一个由 Dictionary 对象中的所有键组成的数组。
|[Remove](#Removemethod)	|从 Dictionary 对象中删除一个指定的键/item对。
|[RemoveAll](#RemoveAllmethod)	|删除 Dictionary 对象中的所有键/item对。

### <a name="Addmethod">Add method</a>

Adds a key and item pair to a Dictionary object.

Syntax
```vba
object.Add key, item
```

The Add method has the following parts:

|Part	|Description
|---|---
|object	|Required. Always the name of a Dictionary object.
|key	|Required. The key associated with the item being added.
|item	|Required. The item associated with the key being added.

> An error occurs if the key already exists.

### <a name="Existsmethod">Exists method</a>

Returns True if a specified key exists in the Dictionary object; False if it does not.

Syntax
```vba
object.Exists (key)
```

The Exists method syntax has these parts:

|Part	|Description
|---|---
|object	|Required. Always the name of a Dictionary object.
|key	|Required. Key value being searched for in the Dictionary object.

### <a name="Itemsmethod">Items method</a>

Returns an array containing all the `items` in a `Dictionary` object.

Syntax
```vba
object.Items
```
The object is always the name of a `Dictionary` object.

Remarks

The following code illustrates use of the Items method:
```vba
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
a = d.Items             'Get the items
For i = 0 To d.Count -1 'Iterate the array
    Print a(i)          'Print item
Next
...
```

### <a name="Keysmethod">Keys method</a>

Returns an array containing all existing `keys` in a `Dictionary` object.

Syntax
```vba
object.Keys
```
The object is always the name of a `Dictionary` object.

Remarks

The following code illustrates use of the Keys method:
```vba
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items.
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
a = d.keys              'Get the keys
For i = 0 To d.Count -1 'Iterate the array
    Print a(i)          'Print key
Next
...
```

### <a name="Removemethod">Remove method</a>

Removes a key/item pair from a `Dictionary` object.

Syntax
```vba
object.Remove (key)
```

The Remove method syntax has these parts:

|Part	|Description
|---|---
|object	|Required. Always the name of a Dictionary object.
|key	|Required. Key associated with the key/item pair that you want to remove from the Dictionary object.

Remarks

> An error occurs if the specified key/item pair does not exist.

The following code illustrates use of the Remove method.

```vba
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
...
d. Remove("b")          'Remove second pair
```

### <a name="RemoveAllmethod">RemoveAll method</a>

The RemoveAll method removes all key, item pairs from a `Dictionary` object.

Syntax
```vba
object.RemoveAll
```
The object is always the name of a `Dictionary` object.

Remarks

The following code illustrates use of the RemoveAll method.
```vba
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
...
a = d.RemoveAll         'Clear the dictionary
```

## 属性

|属性	|说明
|---|---
|[CompareMode](#CompareModeproperty)	|设置或返回用于比较Dictionary对象中的键的比较模式。
|[Count](#Countproperty)	|返回Dictionary对象中的键/item对的数目。
|[Item](#Itemproperty)	|设置或返回Dictionary对象中的item的值。
|[Key](#Keyproperty)	|为Dictionary对象中的现有键值设置新键值。

### <a name="CompareModeproperty">CompareMode property</a>

Sets and returns the comparison mode for comparing string keys in a Dictionary object.

Syntax
```vba
object.CompareMode [ = compare ]
```

The CompareMode property has the following parts:

|Part	|Description
|---|---
|object	|Required. Always the name of a `Dictionary` object.
|compare	|Optional. If provided, compare is a value representing the comparison mode used by functions such as `StrComp`.

Settings

The compare argument can have the following values:

|Constant	|Value	|Description
|---|---|---
|vbUseCompareOption	|-1	|Performs a comparison by using the setting of the `Option Compare` statement.
|vbBinaryCompare	|0	|Performs a binary comparison.
|vbTextCompare	|1	|Performs a textual comparison.
|vbDatabaseCompare	|2	|Microsoft Access only. Performs a comparison based on information in your database.

Remarks

> An error occurs if you try to change the comparison mode of a `Dictionary` object that already contains data.

The `CompareMode` property uses the same values as the compare argument for the `StrComp` function. Values greater than 2 can be used to refer to comparisons by using specific Locale IDs (LCID).

### <a name="Countproperty">Count property</a>

Returns a Long (long integer) containing the number of items in a `collection` or `Dictionary` object. Read-only.

Syntax
```vba
object.Count
```
The object is always the name of one of the items in the Applies To list.

Remarks

The following code illustrates use of the Count property.
```vba
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items.
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
a = d.Keys              'Get the keys
For i = 0 To d.Count -1 'Iterate the array
    Print a(i)          'Print key
Next
...
```

### <a name="Itemproperty">Item property</a>

Sets or returns an item for a specified key in a `Dictionary` object. For collections, returns an item based on the specified key. Read/write.

Syntax
```vba
object.Item (key) [ = newitem ]
```

The Item property has the following parts:

|Part	|Description
|---|---
|object	|Required. Always the name of a `Collection` or `Dictionary` object.
|key	|Required. Key associated with the item being retrieved or added.
|newitem	|Optional. Used for `Dictionary` object only; no application for collections. If provided, newitem is the new value associated with the specified key.

> If key is not found when changing an item, a new key is created with the specified newitem. If key is not found when attempting to return an existing item, a new key is created and its corresponding item is left empty.

### <a name="Keyproperty">Key property</a>

Sets a key in a `Dictionary` object.

Syntax
```vba
object.Key (key) = newkey
```

The Key property has the following parts:

|Part	|Description
|---|---
|object	|Required. Always the name of a `Dictionary` object.
|key	|Required. The key value being changed.
|newkey	|Required. New value that replaces the specified key.

> If key is not found when changing a key, **a run-time error will occur.**
