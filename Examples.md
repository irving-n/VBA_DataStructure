<style>
H4{color:DarkOrange !important;}
</style>

# **Examples**
## **Notes**



> Syntax:
  >> For the sake of shorthand & readability, I'll be using the syntax for Python's data structures, lists, tuples, and dictionaries, to represent the class methods' returns that take the form of the VBA data structures, arrays, collections, and dictionaries respectively.

i.e.,
```Python 
# The elements stored by a VBA array will be of the form:
[1, 2, 3, 4, 5]

# The elements stored by a VBA collection will be of the form:
('Apple', 'Orange', 'Banana', 'Kiwi', 'Mango')

# The {key: item} pairs stored by a VBA Scripting.Dictionary will be of the form:
{'Apple': 20, 'Orange': 3, 'Banana': 5, 'Kiwi': 14, 'Mango': 11}
```
More detail on the analogue in both **Python** & **VBA** for the above can be found in the appendix.
***
<br/>

## **Methods**
## Append
>Adds an element to the end of any of the supported data structures
```VB
Dim todo_list As Variant
Dim task As Variant, task_category As String, task_description As String

Dim housework As Collection
Set housework = New Collection

Dim errands As Scripting.Dictionary
Set errands = New Scripting.Dictionary

todo_list = Array("Clean:Living Room", _
                    "Clean:Kitchen", _
                    "Clean:Room", _
                    "Repair:Leaking Faucet", _ 
                    "Buy:Shoe Rack;1", _
                    "Clean:Window Sills", _
                    "Cook:Lasagna", _
                    "Buy:Detergent;1 Bottle", _
                    "Buy:Milk;2 Cartons", _
                    "Buy:Wool Socks;4 Pairs", _
                    "Return:Library Books;3", _
                    "Return:Faulty Speakers;1")
For Each task In todo_list
    DS.Map(Split(task, ":"), task_category, task_description) ' See Method: Map
    Select Case task_category
        Case "Clean", "Cook", "Repair"
            DS.Append housework, task_description
        Case "Buy", "Return"
            DS.Append errands, Split(task_description, ";")
    End Select
Next task
```
### **housework**
```Python
('Living Room', 'Kitchen', 'Room', 'Leaking Faucet', 'Window Sills', 'Lasagna')
```
### **errands**
```Python
{'Shoe Rack': '1', 'Detergent': '1 Bottle', 'Milk': '2 Cartons', 'Wool Socks': '4 Pairs', 'Library Books': '3', 'Faulty Speakers': '1'}
```
#### - Adds a single element to the supplied data structure

<br/>

***

## Apply
#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| DataStructure_Arr | Variant() | The data structure containing elements to be modified |
| func_name | String | The name of the function being applied |
| arg_pos | Integer | The index position of the argument being supplied on each loop through the data structure |
| other_args | (ParamArray) Variant() | Any of the other arguments required by the function in order |

#### Returns:
##### Example 1: [Vowel Shifting]
> Return

>> Given the following user defined function:
```VB
' This function shifts vowels by 1, relative to their ASC representation
Function shift_vowels(ByVal char As String) As String
    shift_vowels = IIf(InStr("aeiou", LCase(char)) > 0, _
                            Chr(Asc(char) + 1), _
                            char)
End Function
```

```VB
Dim arr As Variant
Dim obfuscated_text As Variant
Dim quote As String
quote = "All that is gold does not glitter, " & _
        "Not all those who wander are lost"
arr = DS.CharacterArray(quote)       '['A', 'l', 'l', ' ', 't', 'h', 'a', 't', ... , 'l', 'o', 's', 't']
obfuscated_text = Join(DS.Apply(arr, "letter_shift", 0), "")
```
> All vowels shifted:
>> Bll thbt js gpld dpfs npt gljttfr, Npt bll thpsf whp wbndfr brf lpst

<br/>

 ##### Example 2: [Simple Math]

```VB
Dim numbers As Variant
Dim output As Variant
arr = Array(2, 4, 6, 8)
output = DS.Apply(arr, "2*", 0)
```

> All elements doubled:
```Python
[4, 8, 12, 16]
```

#### Please note that the feature is fairly limited at the moment since it makes use of the restrictive Eval (Access) / Application.Evaluate (Excel) function
- The mathematical operator and any numbers must come before the element in the array
- The example above creates an expression of the form:
```VB
Application.Evaluate("2*" & "(" & element & ")") 'Where [element] is the control in the for-each-loop governing the output array sourced from the input data structure
```

<br/>

***

## CharacterArray
#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| text | String | The text that needs to be split into separate characters |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> An array with UBound = len(text) - 1, having each character of the provide String as a separate element
```VB
' Comment
Dim book_title as String
Dim all_characters as Variant
book_title = "La Belle Sauvage"
all_characters = DS.CharacterArray(book_title)
```
```Python
['L', 'a', ' ', 'B', 'e', 'l', 'l', 'e', ' ', 'S', 'a', 'u', 'v', 'a', 'g', 'e']
```

<br/>

***
## Convert
#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| DataStructure | Variant(), Collection, Dictionary | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***
## Copy
#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***
## Enumerate
#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***
## Equivalent
#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***

## Exists
#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***

## Fill
#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| container | Variant(), Collection, Dictionary, Integer | The data structure into which elements will be inserted |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

- **container**: Variant() | Collection | Dictionary | Integer
    - The data structure into which elements will be filled
    - OR: The number of elements in the newly created array
- **stuff**: Variant
    - The _"filling"_
- **extra_serving_size**: Integer (Optional)
#### Returns:
> Data structure with type corresponding to the provided data structure, defaulting to Variant() if no data structure is given 
##### 
```VB
' Create a new array & fill with 5 instances of Integer value 1
Dim arr As Variant
arr = DS.Fill(5, 1)
```
> [1, 1, 1, 1, 1]
```VB
' Fill a fixed-size array of upper bound 3 with instances of Integer value 5
Redim arr(3)
DS.Fill arr, 5
```
> [5, 5, 5, 5]

<br/>

***
## Filter

#### Parameters:

| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***
## Flatten
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return

```VB
Dim nested As Variant, flattened As Variant
nested = Array(1, 2, 3, _
                        Array(4, 5, 6), _
                        Array( _
                                Array(7, 8), _
                                9, _
                                Array(10)))
flattened = DS.Flatten(nested)
```
> [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]

<br/>

***
## Homogeneous
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***
## Intersection
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***
## Map
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>

***
## Match
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Maximum
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Merge
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Minimum
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Ones
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Outersection
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Pop
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## PostFixed
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## PreFixed
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Range
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Remove
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Resolve
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Reverse
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Transpose
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/>


***
## Zip
| Variable | Data Type(s) | Description |
| :---: |:--- |:--- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |
| --- | --- | --- |

#### Returns:
> Return
```VB
' Comment
Dim code
```
> Result

<br/><br/><br/>

# Appendix
## Analagous Shorthand Python-VBA
### Array ~ List
```VB
' The elements stored by a VBA array
Array(1, 2, 3, 4, 5)
```

```Python
# Python Analogue:
[1, 2, 3, 4, 5]
```

### Collection ~ Tuple
```VB
' The items comprising the Collection in variable col after executing the following:
Dim col as Collection, arr As Variant, fruit As Variant
Set col = New Collection
arr = Array("Apple", "Orange", "Banana", "Kiwi", "Mango")
For Each fruit In arr
    col.add item:=fruit
Next fruit
```
```Python
# Python Analogue:
("Apple", "Orange", "Banana", "Kiwi", "Mango")
```
### Dictionary ~ Dictionary
```VB
' The {Key:Item} pairs comprising the Scripting.Dictionary in variable dict after executing the following:
Dim dict as Scripting.Dictionary, fruit As Variant, quantities As Variant, i As Integer
Set dict = New Scripting.Dictionary
fruit = Array("Apple", "Orange", "Banana", "Kiwi", "Mango")
quantities = Array(20, 3, 5, 14, 11)
For i = 0 To Ubound(fruit)
    dict.add Key:=fruit(i), Item:=quantities(i)
Next fruit
```
```Python
# Python Analogue:
{'Apple': 20, 'Orange': 3, 'Banana': 5, 'Kiwi': 14, 'Mango': 11}
```
