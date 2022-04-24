VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "DS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'TERMS:
'
'   Iterable: A value-variable or object-variable which holds a set of other values/objects

Private Type tDataStructureStringSettings
InputDelimiter As String
OutputDelimiter As String
End Type


Private Type tDataStructureSettings
Strings As tDataStructureStringSettings
End Type


Private Type tDataStructureProperties
Settings As tDataStructureSettings
DefaultOutputType As String
End Type

Private This As tDataStructureProperties

Private Sub Class_Initialize()
    With This
        With .Settings
            With .Strings
                .InputDelimiter = ";"
                .OutputDelimiter = ";"
            End With
        End With
        .DefaultOutputType = "Variant()"
    End With
End Sub


'PROPERTIES
'Public Property Get Settings() As tDataStructureSettings
'    Settings = This.Settings
'End Property


Public Property Get StringInputDelimiter() As String
StringInputDelimiter = This.Settings.Strings.InputDelimiter
End Property

Public Property Let StringInputDelimiter(ByVal val As String)
This.Settings.Strings.InputDelimiter = val
End Property


Public Property Get StringOutputDelimiter() As String
StringOutputDelimiter = This.Settings.Strings.OutputDelimiter
End Property

Public Property Let StringOutputDelimiter(ByVal val As String)
This.Settings.Strings.OutputDelimiter = val
End Property


'METHODS

Public Function Append(byref body As Variant, byref appendices as variant) As Variant
    if isIterable(body) then
    end if
end function


Public Function Apply(iterable_arr As Variant, funcName As String, elem_arg_idx_pos As Integer, ParamArray args() As Variant) As Variant
    'Note - this function is experimental
    Dim ub As Integer, i As Integer, j As Integer, arg_count As Integer, total_args_count As Integer, xfer_ceiling As Integer
    Dim funcString As String
    Dim rtn As Variant, args_xfer As Variant, val As Variant
    Dim compile_col As Collection, funcPart_col As Collection

    If Not IsArray(iterable_arr) Then
        Apply = Empty
        Exit Function
    End If

    If UBound(args) >= 0 Then
        If UBound(args) = 0 Then 'Fix off by 1 error
            xfer_ceiling = 0
        Else
            xfer_ceiling = UBound(args) + 1
        End If
        total_args_count = 1 + UBound(args) + 1
        ReDim args_xfer(total_args_count - 1)
        args_xfer(elem_arg_idx_pos) = Empty 'This will be replaced in a loop
        j = 0 'position of the input arguments - increment by 1 each time one of them is assigned, pseudo-stack
        '        total_args_count = 1 + UBound(args) + 1
        For i = 0 To (total_args_count - 1)
            If j <= xfer_ceiling Then
                If i <> elem_arg_idx_pos Then
                    On Error Resume Next
                    args_xfer(i) = args(j)
                    If Err.number <> 0 Then
                        Err.Clear
                        Set args_xfer(i) = args(j)
                    End If
                    j = j + 1
                    On Error GoTo 0
                End If
            Else
                '                args_xfer(i) =
            End If
        Next i

        '        ReDim args_xfer(1 + UBound(args))
    Else
        ReDim args_xfer(0)
        total_args_count = 1
    End If

    '    For i = 0 To 5 'can always add more if needed

    ub = UBound(iterable_arr)
    ReDim rtn(ub)

    arg_count = UBound(args)
    Set compile_col = New Collection
    For i = 0 To ub
        On Error Resume Next
        args_xfer(elem_arg_idx_pos) = iterable_arr(i)
        If Err.number <> 0 Then
            Err.Clear
            Set args_xfer(elem_arg_idx_pos) = iterable_arr(i)
        End If
        On Error GoTo 0

        On Error Resume Next
        Select Case total_args_count  ' warning - ugly; it needs to be like this because the arguments supplied can't be filled with a default Missing reference
         Case 1: compile_col.Add ITEM:=Application.run(funcName, args_xfer(0))
         Case 2: compile_col.Add ITEM:=Application.run(funcName, args_xfer(0), args_xfer(1))
         Case 3: compile_col.Add ITEM:=Application.run(funcName, args_xfer(0), args_xfer(1), args_xfer(2))
         Case 4: compile_col.Add ITEM:=Application.run(funcName, args_xfer(0), args_xfer(1), args_xfer(2), args_xfer(3))
         Case 5: compile_col.Add ITEM:=Application.run(funcName, args_xfer(0), args_xfer(1), args_xfer(2), args_xfer(3), args_xfer(4))
         Case 6: compile_col.Add ITEM:=Application.run(funcName, args_xfer(0), args_xfer(1), args_xfer(2), args_xfer(3), args_xfer(4), args_xfer(5))
         Case Else: Debug.Print "Please edit code to add more arguments, follow the same pattern and it should work."
        End Select
        If Err.number <> 0 Then
            funcString = funcName & "("
            Set funcPart_col = New Collection
            For Each val In args_xfer
                If TypeName(val) = "String" Then
                    funcPart_col.Add "'" & val & "'"
                ElseIf IsObject(val) Then
                    'uh oh
                ElseIf IsNull(val) Then
                    val = "Null"
                Else
                    If Not IsEmpty(val) Then funcPart_col.Add val
                    End If
                Next val
                funcString = funcString & Join(CollectionToArray(funcPart_col), ", ") & ")"
                compile_col.Add ITEM:=Eval(funcString)
                Err.Clear
            End If
            On Error GoTo 0
            '        rtn(i) = Application.run(funcName, args(0))
        Next i
        Apply = CollectionToArray(compile_col)
End Function


Public Function Convert(iterable As Variant, Optional ByVal OutputType As String, Optional ConversionOptions As Variant) As Variant
    Dim typeNm As String
    Dim tmpArr As Variant 'temporary array
    Dim rtn As Variant
    Dim col As Collection

    typeNm = TypeName(iterable)
    If Not isIterable(typeNm, True) Then Set Convert = Nothing: Exit Function
        If Not isIterable(OutputType, True) Then Set Convert = Nothing: Exit Function

            Select Case typeNm
             Case "Dictionary"
                tmpArr = Me.Zip(iterable)
             Case "Collection"
                Set col = iterable
                tmpArr = CollectionToArray(col)
             Case Else 'Some form of array
                tmpArr = iterable
            End Select

            If TypeName(tmpArr) = OutputType Then 'no conversion necessary
                rtn = tmpArr
                GoTo Convert_clean_up
            End If

            Select Case OutputType
             Case "Dictionary"
                If IsArray(tmpArr(0)) Then
                    Set rtn = ArrayToDictionary(tmpArr, True)
                Else
                    Set rtn = ArrayToDictionary(tmpArr, False)
                End If
             Case "Collection"
                If IsArray(tmpArr(0)) Then
                    Set rtn = ArrayToCollection(tmpArr, True)
                Else
                    Set rtn = ArrayToCollection(tmpArr, False)
                End If
             Case "Variant()"
                rtn = Me.Apply(tmpArr, "CVar", 0)
             Case "Byte()"
                rtn = Me.Apply(tmpArr, "CByte", 0)
             Case "Integer()"
                rtn = Me.Apply(tmpArr, "CInt", 0)
             Case "Double()"
                rtn = Me.Apply(tmpArr, "CDbl", 0)
             Case "Single()"
                rtn = Me.Apply(tmpArr, "CSng", 0)
             Case "Long()"
                rtn = Me.Apply(tmpArr, "CLng", 0)
             Case "Boolean()"
                rtn = Me.Apply(tmpArr, "CBool", 0)
             Case Else
                'uh oh
            End Select

 Convert_clean_up:
            If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "] in method [Convert] in class DS: " & Err.description: Err.Clear

                Set col = Nothing

                If IsObject(rtn) Then
                    Set Convert = rtn
                Else
                    Convert = rtn
                End If

End Function


Public Function CharacterArray(ByVal text As String) As Variant

    Dim strlen As Integer, i As Integer
    Dim rtn() As Variant
    strlen = Len(text)
    If strlen > 0 Then
        ReDim rtn(strlen - 1)
        For i = 0 To Len(text) - 1
            rtn(i) = Mid(text, i + 1, 1)
        Next i
    End If
    CharacterArray = rtn

End Function


Public Function Reverse(iterable As Variant, Optional ByVal propagate As Boolean = False) As Variant
    Dim typeNm As String
    Dim rtn As Variant, iterable_arr As Variant, col As Collection
    Dim lb As Integer, ub As Integer, i As Integer

    If Not isIterable(iterable) Then
        If IsObject(iterable) Then
            Set Reverse = iterable
        Else
            Reverse = iterable
        End If
        Exit Function
    End If

    'Initialize return container to appropriate type
    typeNm = TypeName(iterable)
    Select Case typeNm
     Case "Dictionary"
        Set rtn = New Scripting.Dictionary
     Case "Collection"
        Set rtn = New Collection
     Case Else
        ReDim rtn(UBound(iterable))
    End Select

    'Convert input into array for easy indexing
    Select Case typeNm
     Case "Dictionary"
        iterable_arr = Zip(iterable) 'Creates nested array containing key-item pairs
     Case "Collection"
        Set col = iterable
        iterable_arr = CollectionToArray(col)
     Case Else
        iterable_arr = iterable 'redundant, but uniform
    End Select

    'Reverse elements
    ub = UBound(iterable_arr)
    lb = LBound(iterable_arr)
    For i = lb To ub
        Select Case typeNm
         Case "Dictionary"
            If propagate And (isIterable(iterable_arr(ub - i)(1))) Then
                rtn.Add Key:=iterable_arr(ub - i)(0), ITEM:=Reverse(iterable_arr(ub - i)(1), propagate:=True)
            Else
                rtn.Add Key:=iterable_arr(ub - i)(0), ITEM:=iterable_arr(ub - i)(1)
            End If
         Case "Collection"
            If propagate And (isIterable(iterable_arr(ub - i))) Then
                rtn.Add ITEM:=Reverse(iterable_arr(ub - i))
            Else
                rtn.Add ITEM:=iterable_arr(ub - i)
            End If
         Case Else
            If propagate And (isIterable(iterable_arr(ub - i))) Then
                rtn(i) = Reverse(iterable_arr(ub - i))
            Else
                rtn(i) = iterable_arr(ub - i)
            End If
        End Select
    Next i

    If IsObject(rtn) Then
        Set Reverse = rtn
    Else
        Reverse = rtn
    End If
End Function


Public Function Exists(element As Variant, search_iterable As Variant, Optional ByVal wildcard_element As Boolean = False, Optional ByVal wildcard_iterable As Boolean = False) As Boolean
    Dim match_found As Boolean, typeNm As String
    Dim search_arr As Variant
    Dim col As Collection
    Dim itm As Variant
    On Error GoTo Exists_clean_up
    match_found = False
    If Not isIterable(search_iterable) Then Exists = match_found: Exit Function

        'Standardize various possible input types for variable search_iterables - normalize as array
        typeNm = TypeName(search_iterable)
        Select Case typeNm
         Case "Dictionary" 'Dictionary has built in method for exists, but this function can expand the method to include wildcards on either the find/element side or the search side
            search_arr = search_iterable.Items
         Case "Collection"
            Set col = search_iterable 'reference assignment is necessary here because byref parameter can't handle a reference to a 2nd reference, to a 3rd reference, etc...
            search_arr = CollectionToArray(col)
         Case Else 'Array
            search_arr = search_iterable
        End Select

        For Each itm In search_arr
            If wildcard_element And (Not IsObject(itm)) And (Not IsObject(element)) Then
                If CStr(itm) Like ("*" & CStr(element) & "*") Then
                    match_found = True
                    Exit For
                End If
            End If
            If wildcard_iterable And (Not IsObject(itm)) And (Not IsObject(element)) Then
                If CStr(element) Like ("*" & CStr(itm) & "*") Then
                    match_found = True
                    Exit For
                End If
            End If
            If IsObject(itm) And IsObject(element) Then
                If itm Is element Then
                    match_found = True
                    Exit For
                End If
            End If
            If (Not IsObject(itm)) And (Not IsObject(element)) Then
                If itm = element Then
                    match_found = True
                    Exit For
                End If
            End If
            If ((IsObject(itm)) And (Not IsObject(element))) _
                Or ((Not IsObject(itm)) And (IsObject(element))) Then
                match_found = False 'explicit, noting here that the search iterable might be composed of BOTH objects and non-objects,
                '   and that the intent is to continue searching through the remainder of the for-loop
            End If
        Next itm

 Exists_clean_up:
        If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "] in method [Exists] in class DS: " & Err.description: Err.Clear
            Set col = Nothing

            Exists = match_found

End Function


Public Function Enumerate(enumerable As Variant, Optional ByVal starting_idx As Variant, Optional ByVal increment As Variant) As Variant
    'I.N. 2/23/22
    'Enumerable types: Array, Collection, Dictionary
    'P/N: Drops keys with dictionaries & collections; adds only the items

    On Error GoTo enumerate_clean_up

    Dim typed_name As String
    Dim val As Variant, i As Integer
    Dim tmp_col As Collection
    Dim rtn As Variant

    Set tmp_col = New Collection

    If IsMissing(starting_idx) Then starting_idx = 0
        i = starting_idx
        typed_name = TypeName(enumerable)

        If Not isIterable(typed_name, True) Then GoTo enumerate_clean_up

            Select Case typed_name
             Case "Variant()", "String()", "Integer()", "Double()", _
                "Single()", "Byte()", "Long()", "Collection":
                For Each val In enumerable
                    tmp_col.Add ITEM:=Array(i, val)
                    i = i + 1
                Next val
             Case "Dictionary":
                For Each val In enumerable.Items
                    tmp_col.Add ITEM:=Array(i, val)
                    i = i + 1
                Next val
             Case Else:
                'incompatible
            End Select

            rtn = CollectionToArray(tmp_col)

 enumerate_clean_up:
            If Err.number <> 0 Then
                Debug.Print "Error [" & Err.number & "]: " & Err.description
                Debug.Print "   Context: " & "Function [enumerate] "
                Err.Clear
                rtn = Array()
            End If

            Set tmp_col = Nothing
            Enumerate = rtn

End Function


Public Function fill(ByRef container As Variant, ByRef stuff As Variant, _
    Optional ByVal extra_serving_size As Long = -1) As Variant
    Dim rtn As Variant
    Dim keyv As Variant
    Dim i As Integer
    Dim init_dim As Variant, add_dim As Variant
    Dim container_ As Variant

    On Error GoTo Fill_clean_up
    If Not isIterable(container, accept_typename:=False) Then
        On Error Resume Next
        If Not IsObject(container) Then
            If IsNumeric(container) Then
                ReDim container_(val(container) - 1)
            Else
                rtn = False
                GoTo Fill_clean_up
            End If
        Else
            rtn = False
            GoTo Fill_clean_up
        End If
        '        rtn = False
        '        GoTo Fill_clean_up
    Else 'iterable
        If IsObject(container) Then
            Set container_ = container
        Else
            container_ = container
        End If
    End If
    On Error GoTo Fill_clean_up

    Select Case TypeName(container_)
     Case "Dictionary"
        If extra_serving_size > 0 Then
            '                container.Add key:=???, item:= stuff
        Else

            For Each keyv In container_.Keys
                On Error Resume Next
                container_(keyv) = stuff
                If Err.number <> 0 Then
                    Err.Clear
                    On Error GoTo Fill_clean_up
                    Set container_(keyv) = stuff
                End If
            Next keyv
        End If

     Case "Collection"
        If extra_serving_size > 0 Then
            For i = 1 To extra_serving_size
                container_.Add stuff
            Next i
        Else
            For i = 1 To container_.count
                If IsObject(stuff) Then
                    Set container_(i) = stuff
                Else
                    container_(i) = stuff
                End If
            Next i
        End If

     Case Else:
        If extra_serving_size > 0 Then
            init_dim = UBound(container_)
            add_dim = init_dim + extra_serving_size
            ReDim Preserve container_(add_dim)
            For i = init_dim + 1 To add_dim
                If IsObject(stuff) Then
                    Set container_(i) = stuff
                Else
                    container_(i) = stuff
                End If
            Next i

        Else
            For i = 0 To UBound(container_)
                If IsObject(stuff) Then
                    Set container_(i) = stuff
                Else
                    container_(i) = stuff
                End If
            Next i
        End If

    End Select

 Fill_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        rtn = False
    Else
        rtn = True
    End If

    If isIterable(container) Then
        'Save changes to byref container and return a boolean
        If IsObject(container_) Then
            Set container = container_
        Else
            container = container_
        End If

    Else
        'Return a filled array if given a value instead of an actual container to fill
        rtn = container_
    End If
    fill = rtn
End Function

Public Function Filter(iterable As Variant, ByVal operator_expression As String, compare_against As Variant, ParamArray other_args() As Variant) As Variant
    On Error GoTo Filter_clean_up
    Dim COLander As Collection
    Dim arr As Variant, val As Variant, thing As Variant
    Dim rtn As Variant
    Dim operand_isIterable As Boolean

    If Not isIterable(iterable, accept_typename:=False) Then
        GoTo Filter_clean_up
    Else
        If TypeName(iterable) = "Dictionary" Then
            'more complicated
        Else
            arr = Me.Convert(iterable, "Variant()")
        End If
    End If

    operand_isIterable = isIterable(compare_against)

    operator_expression = Trim(operator_expression)
    If Me.Exists(LCase(operator_expression), Array("in", "is in", "not in", "is not in", "inside", "between", "out", "outside", "beyond")) Then
        If operand_isIterable Then
            operand_isIterable = False 'Though technically an iterable operator, the intention is not to iterate on each individual element, but to check the operand as a whole
        End If
    End If

    Set COLander = New Collection
    Select Case LCase(operator_expression)
     Case "=", _
        "<>", _
        "is", _
        "is not", "isn't", _
        ">", ">=", _
        "<", "<=", _
        "like", _
        "not like", "liken't", _
        "in", "is in", "not in", "is not in", _
        "inside", "between", _
        "out", "outside", "beyond"
        If operand_isIterable Then
            'Check against many
        Else
            For Each thing In iterable
                If Resolve(thing, LCase(operator_expression), compare_against) Then
                    COLander.Add thing
                Else
                    'Don't add
                End If
            Next thing
        End If
        '
        '        Case "<>"
        '        Case "is", "Is", "IS" 'Objects only
        '
        '        Case "is not", "Is Not", "IS NOT"
        '        Case ">"
        '        Case ">="
        '        Case "<"
        '        Case "<="
        '        Case "like", "Like", "LIKE"
        '        Case "not like", "Not Like", "NOT LIKE"
     Case "like[or]" ', "Like[Or]", "LIKE[OR]"
     Case "like[and]" ', "Like[And]", "LIKE[AND]"
     Case "like[xor]" ', "Like[Xor]", "LIKE[XOR]"
     Case "like[xand]" ', "Like[Xand]", "LIKE[XAND]"
        '        Case "in", "In", "IN"
        '        Case "not in", "Not In", "NOT IN"
        '        Case "between", "inside"
        '        Case "beyond", "out", "outside"
     Case Else
    End Select

    If IsObject(iterable) Then
        Set rtn = Convert(COLander, TypeName(iterable))
    Else
        rtn = Convert(COLander, TypeName(iterable))
    End If

 Filter_clean_up:
    If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "] in method [Filter] in class DS: " & Err.description: Err.Clear
        Set COLander = Nothing
        If IsObject(rtn) Then
            Set Filter = rtn
        Else
            Filter = rtn
        End If
End Function


Public Function Flatten(embedded_iterable As Variant, Optional ByVal OutputType As String = "") As Variant
    Dim recursive_col As Collection
    Dim rtn As Variant
    Dim thing As Variant, validated_iterable As Variant
    Dim rtn_type As String

    If isIterable(embedded_iterable) Then
        If isIterable(OutputType) Then
            rtn_type = OutputType
        Else
            rtn_type = TypeName(embedded_iterable)
        End If
    Else
        Set Flatten = Nothing
        Exit Function
    End If

    Set recursive_col = New Collection

    If TypeName(embedded_iterable) = "Dictionary" Then
        validated_iterable = Zip(embedded_iterable)
    Else
        On Error Resume Next
        validated_iterable = embedded_iterable
        If Err.number <> 0 Then
            Err.Clear
            Set validated_iterable = embedded_iterable
        End If
    End If

    For Each thing In validated_iterable
        If isIterable(thing) Then
            Set recursive_col = Merge(recursive_col, Flatten(thing))
        Else
            recursive_col.Add ITEM:=thing
        End If
    Next thing

    If TypeName(recursive_col) <> rtn_type Then
        On Error Resume Next
        Set rtn = Convert(recursive_col, rtn_type)
        If Err.number <> 0 Then
            Err.Clear
            rtn = Convert(recursive_col, rtn_type)
        End If
    Else
        Set rtn = Convert(recursive_col, rtn_type)
    End If

End Function

Public Function GuiltyByAssociation(associated_iterable As Variant, iterable As Variant, ByVal operator_expression As String, compare_against As Variant, ParamArray other_args() As Variant) As Variant

End Function

Public Function Merge(ParamArray iterables() As Variant) As Variant
    Dim merge_dictionary As Scripting.Dictionary
    Dim submerge As Variant, submerge_arr As Variant, comp As Variant
    Dim copy_iterables As Variant
    Dim output_type As String
    Dim i As Integer, j As Integer
    Dim basekey As String, key_ As String, addkey As String
    Dim dct As Variant, part As Variant

    Set merge_dictionary = New Scripting.Dictionary

    output_type = TypeName(iterables(0))
    If Not isIterable(output_type, accept_typename:=True) Then Exit Function

        For Each submerge In iterables
            i = 1
            If TypeName(submerge) = "Dictionary" And output_type = "Dictionary" Then
                'Splice together keys and items to the main dictionary while retaining keys
                For Each comp In Zip(submerge)
                    j = 0
                    basekey = comp(0)
                    addkey = basekey
                    While merge_dictionary.Exists(addkey)
                        addkey = basekey & CStr(j)
                        j = j + 1
                    Wend
                    merge_dictionary.Add Key:=addkey, ITEM:=comp(1)
                Next comp
            Else
                'Convert everything to an array to add to dictionary
                If TypeName(submerge) = "Dictionary" Then
                    submerge_arr = submerge.Items
                Else
                    submerge_arr = submerge
                End If

                For Each part In submerge_arr
                    merge_dictionary.Add Key:=i, ITEM:=part
                    i = i + 1
                Next part

            End If
        Next submerge

        Select Case output_type
         Case "Dictionary"
            Set Merge = merge_dictionary
         Case "Collection"
            Set Merge = ArrayToCollection(merge_dictionary.Items)
         Case Else
            Merge = merge_dictionary.Items
        End Select

End Function

Public Function Minimum(iterable As Variant, Optional ByVal string_support As Boolean = False) As Variant
    'String support will enable string length comparisons string lengths of provided elements
    Dim val As Variant
    Dim col As Collection
    Dim numeric_composition As Boolean
    Dim string_composition As Boolean
    Dim lowest As Variant
    Dim iterable_copy As Variant

    If Not isIterable(iterable) Then GoTo Minimum_clean_up
        If IterableCount(iterable) < 1 Then GoTo Minimum_clean_up

            Select Case TypeName(iterable)
             Case "Dictionary"
                iterable_copy = iterable.Items
             Case "Collection"
                Set col = iterable
                iterable_copy = CollectionToArray(col)
             Case Else
                iterable_copy = iterable
            End Select

            For Each val In iterable_copy
                If Not IsNumeric(val) Then
                    numeric_composition = False
                    Exit For
                End If
            Next val

            lowest = iterable_copy(0)
            If numeric_composition Then
                For Each val In iterable_copy
                    If val < lowest Then
                        lowest = val
                    End If
                Next val
            End If

 Minimum_clean_up:
            If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "]: " & Err.description: Err.Clear
                Set col = Nothing
                Minimum = lowest
End Function

Public Function Maximum(iterable As Variant) As Variant
    'String support will enable string length comparisons string lengths of provided elements
    Dim val As Variant
    Dim col As Collection
    Dim numeric_composition As Boolean
    Dim string_composition As Boolean
    Dim highest As Variant
    Dim iterable_copy As Variant

    If Not isIterable(iterable) Then GoTo Maximum_clean_up
        If IterableCount(iterable) < 1 Then GoTo Maximum_clean_up

            Select Case TypeName(iterable)
             Case "Dictionary"
                iterable_copy = iterable.Items
             Case "Collection"
                Set col = iterable
                iterable_copy = CollectionToArray(col)
             Case Else
                iterable_copy = iterable
            End Select

            For Each val In iterable_copy
                If Not IsNumeric(val) Then
                    numeric_composition = False
                    Exit For
                End If
            Next val

            highest = iterable_copy(0)
            If numeric_composition Then
                For Each val In iterable_copy
                    If val > highest Then
                        highest = val
                    End If
                Next val
            End If

 Maximum_clean_up:
            If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "]: " & Err.description: Err.Clear
                Set col = Nothing
                Maximum = highest

End Function

Public Function Ones(ParamArray dimensions() As Variant) As Variant
    'Returns an array with size based on the dimensions provided
    'With square brackets denoting an array...
    'Ex: Ones(2, 4, 2) = [ _
    [ _
    [1, 1], _
    [1, 1], _
    [1, 1], _
    [1, 1] _
    ], _
    [ _
    [1, 1], _
    [1, 1], _
    [1, 1], _
    [1, 1] _
    ] _
    ]
    'i.e., 2 Arrays both comprised of 4 arrays, each of which contains 2 1-elements
    '    while me.Fill
End Function

Public Function Pop(ByRef iterable As Variant, Optional ByVal Index As Long) As Variant
    'Return the item indicated by the identifier; mutate the reference to exclude the element denoted by identifier
    Dim rtn As Variant, mutated As Variant
    Dim arr As Variant, comp As Variant
    Dim i As Integer, j As Integer, keyrem As Variant

    If Not isIterable(iterable) Then
        GoTo Pop_clean_up
    End If

    Select Case TypeName(iterable)

     Case "Dictionary"
        If IsMissing(Index) Then Index = 0
            If Index = -1 Then Index = iterable.count - 1 'Last element

                If iterable.count >= Index + 1 Then
                    arr = Me.Zip(iterable.Keys, iterable.Items)
                    keyrem = arr(Index)(0)
                    If IsObject(arr(Index)(1)) Then
                        Set rtn = arr(Index)(1)
                    Else
                        rtn = arr(Index)(1)
                    End If
                    iterable.Remove Key:=keyrem
                End If

             Case "Collection"
                If IsMissing(Index) Then Index = 1
                    If Index = -1 Then Index = iterable.count
                        If Index = 0 Then Index = 1 'Collections are base 1
                            If iterable.count >= Index Then
                                If IsObject(iterable(Index)) Then
                                    Set rtn = iterable(Index)
                                Else
                                    rtn = iterable(Index)
                                End If
                            End If

                            If iterable.count > 0 Then
                                Set iterable = Me.Remove(iterable, Index - 1) 'Method [Remove] uses index base 0
                            End If

                         Case Else 'Array
                            If IsMissing(Index) Then Index = 0
                                If Index = -1 Then Index = UBound(iterable)
                                    If UBound(iterable) = 0 Then
                                        rtn = iterable(0)
                                        iterable = Array()
                                        GoTo Pop_clean_up
                                    End If
                                    ReDim mutated(UBound(iterable) - 1)
                                    rtn = iterable(Index)
                                    j = 0 'index of last addition to mutated array, starting from 0
                                    For i = 0 To UBound(iterable)
                                        If i <> Index Then
                                            mutated(j) = iterable(i)
                                            j = j + 1
                                        End If
                                    Next i
                                    If j <> i - 1 Then
                                        'uh oh
                                        'don't overwrite
                                    Else
                                        iterable = mutated
                                    End If
                                End Select

 Pop_clean_up:
                                If Err.number <> 0 Then
                                    Debug.Print "Error [" & Err.number & "]: " & Err.description
                                    Err.Clear
                                End If

                                If IsObject(rtn) Then
                                    Set Pop = rtn
                                Else
                                    Pop = rtn
                                End If

End Function

Public Function Remove(iterable As Variant, Optional ByVal Index As Integer = 0) As Variant
    Dim col As Collection
    Dim i As Integer
    Dim arr As Variant, rtn As Variant
    Dim rkey As String

    If Not isIterable(iterable) Then
        On Error Resume Next
        Remove = iterable
        If Err.number <> 0 Then Err.Clear: Set Remove = iterable
            Exit Function
        End If

        Select Case TypeName(iterable)
         Case "Dictionary"
            rkey = iterable.Keys(Index)
            Set rtn = iterable.Remove(rkey)
         Case "Collection"
            Set col = iterable
            arr = CollectionToArray(col)
            Set rtn = ArrayToCollection(Me.Remove(arr, Index))
         Case Else
            Set col = New Collection

            For i = 0 To UBound(iterable)
                If i <> Index Then col.Add iterable(i)
                Next i
                rtn = CollectionToArray(col)
            End Select

            If IsObject(rtn) Then
                Set Remove = rtn
            Else
                Remove = rtn
            End If
End Function

Public Function Transpose(ParamArray DataStructures() As Variant) As Variant

End Function

Public Function Zip(ParamArray iterables() As Variant) As Variant
    On Error GoTo Zip_clean_up
    Dim rtn As Variant, copy_iterables As Variant, iterable As Variant, thing As Variant
    Dim output_col As Collection, tcol As Collection
    Dim i As Integer, j As Integer, thisCount As Integer, lastCount As Integer, itCount As Integer

    'Make sure all parameters are the same count/size
    lastCount = -2 'arbitrary, but not -1 since ubound of empty array is -1
    For Each iterable In iterables
        thisCount = IterableCount(iterable)

        If lastCount = -2 Then
            lastCount = thisCount
        End If

        If thisCount <> lastCount Then
            'Incompatible
            Zip = Empty
            Exit Function
        End If
    Next iterable
    itCount = thisCount 'Use later to output and size the return array to the span (itCount) of the input iterables

    'Special case for standalone dictionaries; dictionaries are compacted and returned
    If UBound(iterables) = 0 Then
        If TypeName(iterables(0)) = "Dictionary" Then
            With iterables(0)
                rtn = zipArray_(.Keys, .Items)
            End With
            Zip = rtn
            Exit Function
        End If
    End If

    ReDim copy_iterables(UBound(iterables))
    For i = 0 To UBound(iterables)
        On Error Resume Next
        copy_iterables(i) = iterables(i)
        If Err.number <> 0 Then
            Err.Clear
            If TypeName(iterables(i)) = "Collection" Then
                Set copy_iterables(i) = iterables(i)
            ElseIf TypeName(iterables(i)) = "Dictionary" Then
                copy_iterables(i) = Zip(iterables(i))
            Else
                Debug.Print "Error in Zip (not a standard raised error): the indexed retrieval errored out, and the object passed doesn't qualify as a Collection or Dictionary"
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next i

    On Error GoTo Zip_clean_up
    Set output_col = New Collection

    ReDim rtn(itCount - 1)

    For j = 0 To itCount - 1
        Set tcol = New Collection
        For Each thing In copy_iterables
            Select Case TypeName(thing)
             Case "Collection"
                tcol.Add ITEM:=thing(j + 1)
             Case "Dictionary" 'uh oh
                tcol.Add ITEM:=thing(thing.Keys(j))
             Case Else
                tcol.Add ITEM:=thing(j)
            End Select
            '            Stop
        Next thing
        rtn(j) = CollectionToArray(tcol)
    Next j
 Zip_clean_up:
    If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "]: " & Err.description
        Zip = rtn
End Function


'PRIVATE AUXILIARY METHODS
Private Function IterableCount(ByRef iterable As Variant) As Integer
    On Error GoTo IterableCount_clean_up
    Dim i As Integer, val As Variant
    i = 0
    For Each val In iterable
        i = i + 1
    Next val
 IterableCount_clean_up:
    If Err.number <> 0 Then Err.Clear
        '    If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "]: " & Err.description
        IterableCount = i
End Function


Private Function IterableMaxSpan(ByRef iterable As Variant) As Long

End Function


Private Function IterableToArray(iterable As Variant) As Integer

End Function


Private Function isIterable(iterable As Variant, Optional ByVal accept_typename As Boolean = True) As Boolean
    Dim typnm As String
    Dim rtn As Boolean: rtn = False
    'Note:
    '       Accepts Syntax:
    '1.           isIterable("Variant()")
    '2.          isIterable(object)
    If accept_typename And TypeName(iterable) = "String" Then
        typnm = iterable
    Else
        typnm = TypeName(iterable)
    End If
    Select Case typnm 'Accepts string name of iterable type as well as the iterable itself
     Case "Variant()", "Long()", _
        "Single()", "Byte()", _
        "Integer()", "Double()", _
        "String()", "Boolean()", _
        "Collection", "Dictionary"
        rtn = True
     Case Else
        rtn = False
    End Select
    isIterable = rtn
End Function


Private Function CollectionToArray(iterable_col As Collection) As Variant
    Dim arr As Variant
    Dim i As Integer
    If iterable_col.count > 0 Then
        ReDim arr(iterable_col.count - 1)
        For i = 0 To UBound(arr)
            If IsObject(iterable_col(i + 1)) Then
                Set arr(i) = iterable_col(i + 1)
            Else
                arr(i) = iterable_col(i + 1)
            End If
        Next i
    Else
        arr = Array()
    End If
    CollectionToArray = arr
End Function


Private Function DictionaryToArray(iterable_dictionary As Scripting.Dictionary) As Variant
    Dim comp As Variant, keyarr As Variant, itemarr As Variant
    Dim arr() As Variant
    Dim i As Integer
    If iterable_dictionary.count > 0 Then
        ReDim arr(iterable_dictionary.count - 1)
        keyarr = iterable_dictionary.Keys
        itemarr = iterable_dictionary.Items
        For i = 0 To (iterable_dictionary.count - 1)
            arr(i) = Array(keyarr(i), itemarr(i))
        Next i
    End If
    DictionaryToArray = arr
End Function


Private Function ArrayToCollection(iterable As Variant, Optional ByVal populateKeys As Boolean = True) As Collection
    Dim rtn_col As Collection
    Dim val As Variant
    Set rtn_col = New Collection

    If IsArray(iterable) Then
        On Error Resume Next
        If UBound(iterable) >= 0 Then
            If Err.number <> 0 Then
                'No items to add
            Else
                For Each val In iterable
                    rtn_col.Add val
                Next val
            End If
        End If
    End If

    Set ArrayToCollection = rtn_col
End Function


Private Function ArrayToDictionary(iterable As Variant, Optional ByVal populateKeys As Boolean = True) As Scripting.Dictionary
    Dim rtn_dictionary As Scripting.Dictionary
    Dim nextkey As String, basekey As String
    Dim val As Variant, comp As Variant
    Dim keyidx_suffix As Integer, i As Integer
    Set rtn_dictionary = New Scripting.Dictionary

    If IsArray(iterable) Then
        If IsArray(iterable(0)) Then
            If populateKeys Then
                For Each comp In iterable
                    keyidx_suffix = 0
                    basekey = CStr(comp(0))
                    nextkey = basekey
                    While rtn_dictionary.Exists(nextkey)
                        nextkey = basekey & CStr(keyidx_suffix)
                        keyidx_suffix = keyidx_suffix + 1
                    Wend
                    rtn_dictionary.Add Key:=nextkey, ITEM:=Remove(comp, Index:=0)
                Next comp
            Else 'ignore keys, use regular indexing
                i = 0
                For Each val In iterable
                    rtn_dictionary.Add Key:=i, ITEM:=val
                    i = i + 1
                Next val
            End If 'populate keys
        Else
            i = 0
            For Each val In iterable
                rtn_dictionary.Add Key:=i, ITEM:=val
                i = i + 1
            Next val
        End If 'isArray
    Else
        'input isn't an array - don't do anything
    End If '

 ArrayToDictionary_clean_up:
    If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "]: " & Err.description

        Set ArrayToDictionary = rtn_dictionary

End Function

Private Function zipArray_(ParamArray args() As Variant) As Variant()
    'Released I.N. - 2/23/22

    Dim arg As Variant, sub_arg As Variant
    Dim ub_arg As Long, ub_args As Integer, lb_arg As Long, lb_args As Integer
    Dim sz_of_sub_rtn As Integer, sz_of_rtn As Long
    Dim rtn() As Variant, sub_rtn() As Variant
    Dim i As Long, j As Integer
    Dim valid_args As Boolean: valid_args = True

    'Example usage:
    'Input:
    '   A:= array(1, 2, 3)
    '   B:= array(4, 5, 6)
    '   C:= array(7, 8, 9)

    'Output:
    '   Return= array(array(1, 4, 7), array(2, 5, 8), array(3, 6, 9))

    'Check first if they're all the same size
    For Each arg In args
        valid_args = IsArray(arg)
        If valid_args Then If UBound(args) = -1 Then valid_args = False Else valid_args = True
            If Not valid_args Then Debug.Print "Argument is not an array.": zipArray_ = Array(Null): Exit Function
            Next arg

            'Number of params provided in array
            ub_args = UBound(args)
            lb_args = LBound(args)
            sz_of_sub_rtn = ub_args - lb_args

            'Check if sub-arrays are the same size
            ub_arg = UBound(args(0))
            lb_arg = LBound(args(0))
            sz_of_rtn = ub_arg - lb_arg

            For Each arg In args
                If UBound(arg) <> ub_arg Or LBound(arg) <> lb_arg Then valid_args = False
                    If Not valid_args Then
                        Debug.Print "Array bounds do not match."
                        zipArray_ = Array(Null)
                        Exit Function
                    End If
                Next arg

                'Arguments are valid - continue...

                'Resize sub arrays and array that will be returned
                ReDim rtn(sz_of_rtn)
                ReDim sub_rtn(sz_of_sub_rtn)

                'Combine the 1st element of each array, then the 2nd... etc..
                For i = 0 To sz_of_rtn
                    For j = 0 To sz_of_sub_rtn

                        On Error Resume Next
                        'Try value assignment
                        sub_rtn(j) = args(j)(i)
                        If Err.number <> 0 Then
                            Err.Clear
                            'Try reference object
                            Set sub_rtn(j) = args(j)(i)
                        End If
                        'Finally
                        On Error GoTo 0

                    Next j

                    'Add the combined array to the stack of arrays
                    rtn(i) = sub_rtn
                Next i

                'Return the result
                zipArray_ = rtn
End Function

Private Function Resolve(left_side As Variant, operator As String, right_side As Variant) As Boolean
    On Error GoTo Resolve_clean_up
    Dim vali As Variant
    Dim lowercase_operator As String, mini_ As Variant, maxi_ As Variant
    lowercase_operator = LCase(operator)
    Dim rtn As Boolean

    rtn = False

    Select Case lowercase_operator
     Case "="
        rtn = (left_side = right_side)
     Case "<>"
        rtn = (left_side <> right_side)
     Case "is", "Is", "IS" 'Objects only
        If IsObject(left_side) And IsObject(right_side) Then
            rtn = (left_side Is right_side)
        End If
     Case "is not"
        If IsObject(left_side) And IsObject(right_side) Then
            rtn = (left_side Is Not right_side)
        End If
     Case ">"
        If IsNumeric(left_side) And IsNumeric(right_side) Then
            rtn = (left_side > right_side)
        End If
     Case ">="
        If IsNumeric(left_side) And IsNumeric(right_side) Then
            rtn = (left_side >= right_side)
        End If
     Case "<"
        If IsNumeric(left_side) And IsNumeric(right_side) Then
            rtn = (left_side < right_side)
        End If
     Case "<="
        If IsNumeric(left_side) And IsNumeric(right_side) Then
            rtn = (left_side <= right_side)
        End If
     Case "in", "is in"
        rtn = Exists(left_side, right_side)
     Case "like"
        rtn = (CStr(left_side) Like ("*" & CStr(right_side) & "*"))
     Case "not like"
        rtn = (Not (CStr(left_side) Like ("*" & CStr(right_side) & "*")))
     Case "between"
        If Not IsArray(right_side) Then GoTo Resolve_clean_up
            If UBound(right_side) <> 1 Then GoTo Resolve_clean_up
                rtn = ((left_side >= Minimum(right_side)) And (left_side <= Maximum(right_side)))
             Case "outside"
                If Not IsArray(right_side) Then GoTo Resolve_clean_up
                    If UBound(right_side) <> 1 Then GoTo Resolve_clean_up
                        rtn = ((left_side < Minimum(right_side)) Or (left_side > Maximum(right_side)))
                     Case Else
                    End Select
 Resolve_clean_up:

                    Resolve = rtn
End Function


Private Function hasUniformComposition(arr As Variant) As Boolean
    Dim rtn As Boolean
    Dim val As Variant
    Dim tn As String

    rtn = True 'uniform until proven otherwise
    If IsArray(arr) Then
        tn = TypeName(arr(0))
        For Each val In arr
            If TypeName(val) <> tn Then
                rtn = False
                Exit For
            End If
        Next val
    Else
        rtn = False
    End If
    hasUniformComposition = rtn
End Function


