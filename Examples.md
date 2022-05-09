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
| OutputType | String | The name of the data type |
| ConversionOptions | Variant | No current implementation |
| keys | Variant() | For conversion to dictionaries and collections, array of keys with the same number of elements as the DataStructure |

#### Returns:

The data structure converted into whatever format was specified

<br/>

#### Example 1: Basic Conversions

```VB
Dim dict As Scripting.Dictionary
Dim ingredients() As String
Dim keys As Variant
Dim in_stock() As Boolean

Set dict = New Scripting.Dictionary


```



#### Example 2: Array(s) to dictionary of collections

#### Napkin-Math, except done in bulk.
#### **A quick cost-performance analysis: maximum point load under fixed cantilever loading conditions of various metals**



*The following example uses the **Convert** method to associate several pieces of linked data.*

Material property sources: 
- https://www.engineeringtoolbox.com/young-modulus-d_417.html
- https://www.mcmaster.com
- https://www.aerospacemetals.com/aluminum-distributor.html


> Length
>> L = 3 ft = 36 in

> Cross Section 
>> a = b = 0.25 in (Square)

> Distance to Neutral Axis
>> y = 0.5*a
>>> = 0.125 in

> Area Moment of Inertia 
>> I = a<sup>4</sup>/12
>>> = 3.26e<sup>-4</sup> in<sup>4</sup>

<br/><br/>
> Max Force, F<sub>max</sub>, on a cantilever before yielding:

> F = σ I / y L


```VB
Function force_at_yield(ByVal yield_strength As Double, ByVal area_mom_inert As Double, ByVal neutral_axis_dist As Double, ByVal length As Double) As Double
    force_at_yield = (yield_strength * area_mom_inert) / (neutral_axis_dist * length)
End Function
```

```VB
Function deflection_at_yield(ByVal force As Double, ByVal length As Double, ByVal elast_mod As Double, ByVal area_mom_inert As Double) As Double
    deflection_at_yield = (force * (length ^ 3)) / (3 * elast_mod * area_mom_inert)
End Function
```

```VB
Dim l1 As Double, ArMoIn As Double, y_neut As Double, dataset As Variant
Dim mat_name As String, md As Variant, max_force As Double, max_defl As Double
Dim materials As Variant
Dim elastic_moduli As Variant
Dim yield_strengths As Variant
Dim mcm_ids As Variant
Dim prices_per_lineal_ft As Variant
Dim material_properties As Variant
Dim header_keys As Variant
Dim property_set As Variant, props_col As Collection, prop_dictionary As Scripting.Dictionary

' SHAPE -----------------------------------
'Length
l1 = 1 * 12 ' in.
'Area Moment of Inertia
ArMoIn = 0.000326 'in ^ 4
'Distance to Neutral Axis
y_neut = 0.125 'in
' -----------------------------------------

' MATERIAL PROPERTIES ---------------------
'name
materials = Array("Aluminum:Anodized Multipurpose 6061", "Aluminum:Architectural 6063", "Aluminum:High-Strength 2024", "Aluminum:Easy-to-Machine 2011", "Low-Carbon Steel Bar 1018", "Ultra-Machinable 12L14 Carbon Steel Bars", "A2 Tool Steel")
'modulus of elasticity
elastic_moduli = Array(10000, 10000, 10600, 10150, 29700, 29000, 27500) 'ksi
elastic_moduli = DS.Apply(elastic_moduli, "1000*", 0) 'ksi -> psi
'yield strength
yield_strengths = Array(35000, 16000, 47000, 38000, 54000, 60000, 51000) 'psi
'McMaster
mcm_ids = Array("6023K35", "89755K69", "86895K81", "3031N2", "9143K13", "6547K112", "9019K95")
'price
prices_per_lineal_ft = Array(24.99 / 3, 8.39 / 8, 57.29 / 6, 17.94 / 6, 9.59 / 6, 42.97 / 6, 123.21 / 6) '$/ft
' -----------------------------------------

header_keys = Array("name", "modulus of elasticity", "yield strength", "McMaster", "price")
material_properties = DS.Zip(materials, elastic_moduli, yield_strengths, mcm_ids, prices_per_lineal_ft)
Set prop_dictionary = New Scripting.Dictionary

For Each property_set In material_properties
    Set props_col = DS.Convert(property_set, "Collection", keys:=header_keys)
    prop_dictionary.Add Key:=props_col("name"), Item:=props_col
Next property_set

Debug.Print "Analysis Results for .25x.25 in^2 square bar, length: " & l1 & "in"
Debug.Print "Loading configuration: Point-load, cantilever"
Debug.Print
For Each dataset In DS.Zip(prop_dictionary)
    mat_name = dataset(0) 'material name
    Set md = dataset(1) 'material dataset
    max_force = force_at_yield(md("yield strength"), ArMoIn, y_neut, l1)
    max_defl = deflection_at_yield(max_force, l1, md("modulus of elasticity"), ArMoIn)
    
'    Debug.Print "Analysis Results for .25x.25 in^2 [" & mat_name & "] square bar, length: " & l1 & "in"
    Debug.Print "Material: [" & mat_name & "]"
    Debug.Print Tab(10); "Yield occurs under [" & Format(CStr(max_force), "#.##") & "] lbs after deflecting [" & Format(CStr(max_defl), "#.###") & "] inches."
'    Debug.Print Tab(15); "At which point, it will have deflected [" & Format(CStr(max_defl), "#.###") & "] inches."
    Debug.Print Tab(5); "Price: " & Format(CStr(md("price") * (l1 / 12)), "$#.##")
    Debug.Print
Next dataset
```

> Analysis Results for .25x.25 in^2 square bar, length: 12in

> **Loading configuration**: Point-load, cantilever

<br/>

> **Material**: [Aluminum:Anodized Multipurpose 6061]
>> Yield occurs under [7.61] lbs after deflecting [1.344] inches.

>> Price: $8.33

**Material**: [Aluminum:Architectural 6063]
>> Yield occurs under [3.48] lbs after deflecting [.614] inches.

>> Price: $1.05

**Material**: [Aluminum:High-Strength 2024]
>> Yield occurs under [10.21] lbs after deflecting [1.703] inches.

>> Price: $9.55

**Material**: [Aluminum:Easy-to-Machine 2011]
>> Yield occurs under [8.26] lbs after deflecting [1.438] inches.

>> Price: $2.99

**Material**: [Low-Carbon Steel Bar 1018]
>> Yield occurs under [11.74] lbs after deflecting [.698] inches.

>> Price: $1.6

**Material**: [Ultra-Machinable 12L14 Carbon Steel Bars]
>> Yield occurs under [13.04] lbs after deflecting [.794] inches.

>> Price: $7.16

**Material**: [A2 Tool Steel]
>> Yield occurs under [11.08] lbs after deflecting [.712] inches.

>> Price: $20.54

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
