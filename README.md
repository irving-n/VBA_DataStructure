# VBA_DataStructure

## What is this project?
This project aims to simplify commonly used data structures in VBA, i.e., Arrays, Collections, Dictionaries, by both replicating existing methods and providing additional methods - all encapsulated under a single easy-to-use class.

## What kind of methods are available?
This project is still in development - These are still subject to change!    
Append, Apply, Copy, Enumerate, Fill, Filter, Having, Map, Maximum, Merge, Minimum, Ones, Reverse, Sort, Transpose, Zip, and more...

## What are the system requirements & dependencies?
### Libaries:
Microsoft Scripting Runtime[^1]
### OS[^2]:
    Windows 7
    Windows 10

##### Footnotes:
[^1]: 
    Those who want late-bound behavior may replace variable declarations of Scripting.Dictionary with Object.  
    E.g.:
    ```VB
    Dim my_dictionary As Scripting.Dictionary
    ```
    Becomes
    ```VB
    Dim my_dictionary As Object
    ```
[^2]:  
    MAC users may encounter issues regarding the usage of the **Scripting.Dictionary** object derived from the Microsoft Scripting Runtime library.
