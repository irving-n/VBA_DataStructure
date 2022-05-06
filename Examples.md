# Examples
## Notes 



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

## Methods
### Append
>Adds an element to the end of any of the supported data structures
```VB

```
#### - Adds a single element to the supplied data structure
***
### Apply
#### - Applies the provided function to all elements in the supplied data structure; returns 
***
### CharacterArray
***
### Convert
***
### Copy
***
### Enumerate
***
### Equivalent
***
### Exists
***
### Fill
```VB
Dim arr As Variant
arr = DS.Fill(5, 1)
```
> arr:
> (1, 1, 1, 1, 1)
```VB
Redim arr(3)
DS.Fill(arr, 5)
```
> arr:
> (5, 5, 5, 5)
***
### Filter
***
### Flatten
***
### Homogeneous
***
### Intersection
***
### Map
***
### Match
***
### Maximum
***
### Merge
***
### Minimum
***
### Ones
***
### Outersection
***
### Pop
***
### PostFixed
***
### PreFixed
***
### Range
***
### Remove
***
### Resolve
***
### Reverse
***
### Transpose
***
### Zip



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
