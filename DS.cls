VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#Const conApp = "Microsoft Excel"
#Const conApp = "Microsoft Access"
'TERMS:
'
'   DataStructure: A value-variable or object-variable which holds a set of other values/objects

Public Enum eDataStructureFilterMode
    eFilterTrap
    eFilterOut
End Enum

Private Sub Class_Initialize()
    'Nothing to initialize - This is only an interface - type class; it doesn't store anything internally
End Sub

Private Sub Class_Terminate()
    'Nothing to clean up
End Sub

'PROPERTIES


'METHODS

Public Function Append(ByRef body As Variant, ByRef appendix As Variant) As Variant
    'Syntax:
    '   new_arr = DS.Append(old_arr, new_elem)
    
    Dim output_type As String
    Dim arr As Variant, pend As Variant
    Dim current_ubound As Long, new_ubound As Long, k As Long
    Dim newkey As String

    
On Error GoTo Append_clean_up

    If isDataStructure(body) Then
        output_type = TypeName(body)
        Select Case output_type
            Case "Dictionary"
                k = 0
                newkey = CStr(k)
                While body.Exists(newkey)
                    k = k + 1
                    newkey = CStr(k)
                Wend
                body.Add Key:=newkey, ITEM:=appendix
                    
            Case "Collection"
                body.Add appendix
            Case Else
                On Error Resume Next
                
                current_ubound = UBound(body)
                new_ubound = current_ubound + 1
                
                If Err.number <> 0 Then Err.Clear
                ReDim Preserve body(new_ubound)
                
                If IsObject(appendix) Then
                    Set body(new_ubound) = appendix
                  Else
                    body(new_ubound) = appendix
                End If
                
        End Select
    
    End If
    
Append_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Err.Clear
    End If
    
    If IsObject(body) Then
        Set Append = body
      Else
        Append = body
    End If
End Function


Public Function Apply(DataStructure_arr As Variant, funcName As String, elem_arg_idx_pos As Integer, ParamArray args() As Variant) As Variant
'Syntax:
'   output = DS.Apply(old_arr, '<some function name as a string>', index_pos, remaining arguments)

'elem_arg_idx_pos: (BASE 0) Refers to the placement of arguments from the data structure into the desired function

'For example: function InStr:
'arr_of_strings = Array("aaaa FIND_ME aaaa", "aa FIND_ME aa", "aaa FIND_ME aaa", "aaaaaaaa")

'To get the string position of FIND_ME within each of the strings contained in this array:
'string_positions = DS.Apply(arr_of_strings, "InStr", 1, 1, "FIND_ME")
'   Outputs: (6, 4, 5, 0)


    Dim ub As Integer, i As Integer, j As Integer, arg_count As Integer, total_args_count As Integer, xfer_ceiling As Integer
    Dim funcString As String
    Dim rtn As Variant, args_xfer As Variant, val As Variant
    Dim delm As String
    Dim compile_col As Collection, funcPart_col As Collection
    Dim this_app As Application
    
    If Not IsArray(DataStructure_arr) Then
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
    
    ub = UBound(DataStructure_arr)
    ReDim rtn(ub)
    
    arg_count = UBound(args)
    Set compile_col = New Collection
    For i = 0 To ub
        On Error Resume Next
        args_xfer(elem_arg_idx_pos) = DataStructure_arr(i)
        If Err.number <> 0 Then
            Err.Clear
            Set args_xfer(elem_arg_idx_pos) = DataStructure_arr(i)
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
            If DS.Exists(funcName, Array("Join", "Split")) Then
                Set funcPart_col = New Collection
                For Each val In args_xfer
                    funcPart_col.Add val
                Next val
                
                If funcPart_col.count > 1 Then delm = funcPart_col(2) Else delm = " "
                
                If funcName = "Join" Then
                    If delm = vbCrLf Then
                        compile_col.Add ITEM:=Join(funcPart_col(1), vbCrLf)
                      Else
                        compile_col.Add ITEM:=Join(funcPart_col(1), delm)
                    End If
                  Else
                    If delm = vbCrLf Then
                        compile_col.Add ITEM:=split(funcPart_col(1), vbCrLf)
                      Else
                        compile_col.Add ITEM:=split(funcPart_col(1), delm)
                    End If
                End If
              Else
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
                'Set this_app = Application
                
                #If conApp = "Microsoft Excel" Then
                    compile_col.Add ITEM:=Application.Evaluate(funcString)
                #Else
                    compile_col.Add ITEM:=Eval(funcString)
    '                Debug.Print "Not Excel!!!"
                #End If
            End If

            
            Err.Clear
        End If
        On Error GoTo 0
'        rtn(i) = Application.run(funcName, args(0))
    Next i
    Apply = CollectionToArray(compile_col)
End Function


Public Function CharacterArray(ByVal text As String) As Variant
'Syntax
'   arr = DS.CharacterArray("abcdefg")
'>> ("a", "b", "c", "d", "e", "f", "g")
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


Public Function Convert(DataStructure As Variant, ByVal OutputType As String, Optional ConversionOptions As Variant, Optional keys As Variant) As Variant
    Dim typeNm As String
    Dim tmpArr As Variant 'temporary array
    Dim rtn As Variant
    Dim col As Collection
On Error GoTo Convert_clean_up
    typeNm = TypeName(DataStructure)
    If Not isDataStructure(typeNm, True) Then Set Convert = Nothing: Exit Function
    If Not isDataStructure(OutputType, True) Then Set Convert = Nothing: Exit Function
    
    Select Case typeNm
        Case "Dictionary"
            tmpArr = Me.Zip(DataStructure)
        Case "Collection"
            Set col = DataStructure
            tmpArr = CollectionToArray(col)
        Case Else 'Some form of array
            tmpArr = DataStructure
    End Select
    
    If TypeName(tmpArr) = OutputType Then 'no conversion necessary
        rtn = tmpArr
        GoTo Convert_clean_up
    End If
    
    Select Case OutputType
        Case "Dictionary"
            If Not IsMissing(keys) Then
                Set rtn = ArrayToDictionary(DS.Zip(keys, tmpArr), True)
              ElseIf IsArray(tmpArr(0)) Then
                Set rtn = ArrayToDictionary(tmpArr, True)
              Else
                Set rtn = ArrayToDictionary(tmpArr, False)
            End If
        Case "Collection"
            On Error Resume Next
            If (Not IsMissing(keys)) Then
                Set rtn = ArrayToCollection(tmpArr, True, keys:=keys)
              ElseIf IsArray(tmpArr(0)) Then
                Set rtn = ArrayToCollection(tmpArr, IIf(Err.number = 0, True, False))
              Else
                Set rtn = ArrayToCollection(tmpArr, False)
            End If
            On Error GoTo Convert_clean_up
        Case "Variant()"
            rtn = Me.Apply(tmpArr, "CVar", 0)
        Case "String()"
            rtn = Me.Apply(tmpArr, "CStr", 0)
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
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "] in method [Convert] in class DS: " & Err.description: Err.Clear
        Resume Next
    End If
    Set col = Nothing
    
    If IsObject(rtn) Then
        Set Convert = rtn
      Else
        Convert = rtn
    End If
    
End Function


Public Function Copy(DataStructure As Variant, Optional ByVal failOnNestedObjects As Boolean = False) As Variant
    'Considering renaming or having aliases for this method: Clone, Duplicate - might better convey the doubling of memory behind the scenes
    '
    'Creates and returns a new copy without references to original
    'Will not handle nested objects for now; possible to handle nested DataStructures recursively in the future
    Dim rtn As Variant
    Dim thing As Variant, comp As Variant, span As Variant
    Dim curkey As String, curitem As Variant
    Dim i As Long, lb As Long, ub As Long
    
    On Error GoTo Copy_clean_up
    If isDataStructure(DataStructure) Then
        Select Case TypeName(DataStructure)
            Case "Dictionary"
                Set rtn = New Scripting.Dictionary
                For Each comp In Zip(DataStructure.keys, DataStructure.Items)
                    curkey = comp(0)
                    If IsObject(comp(1)) Then
                        If failOnNestedObjects Then Err.Raise 450
                        Set curitem = DS.Copy(comp(1)) 'Warning: recursive
                      Else
                        curitem = comp(1)
                    End If
                    rtn.Add Key:=curkey, ITEM:=curitem
                Next comp
                
            Case "Collection"
                Set rtn = New Collection
                For Each thing In DataStructure
                    If IsObject(thing) Then
                        If failOnNestedObjects Then
                            Err.Raise 450
                          'no need to have an Else clause as the "add" method will accept both values & references
                        End If
                    End If
                    rtn.Add thing
                Next thing
                
            Case Else
                'Catch uninitialized arrays
                span = -1
                
                On Error Resume Next
                lb = LBound(DataStructure)
                ub = UBound(DataStructure)
                
                span = ub - lb
                
                If Err.number = 0 Then
                    ReDim rtn(LBound(DataStructure) To UBound(DataStructure))
                  Else
                    GoTo Copy_clean_up
                End If
                
                On Error GoTo Copy_clean_up
                
                For i = lb To ub
                    If IsObject(DataStructure(i)) Then
                        If failOnNestedObjects Then
                            Err.Raise 450
                          Else
                            Set rtn(i) = Me.Copy(DataStructure(i)) 'Warning: recursive
                        End If
                      Else
                        If IsArray(DataStructure(i)) Then
                            rtn(i) = DS.Copy(DataStructure(i))
                          Else
                            rtn(i) = DataStructure(i)
                        End If
                    End If
                Next i
        End Select
    End If
                
Copy_clean_up:

    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Err.Clear
    End If
    
    If IsObject(rtn) Then
        Set Copy = rtn
      Else
        Copy = rtn
    End If
    
End Function


Public Function Range(ByVal start_value As Variant, ByVal end_value As Variant, Optional ByVal step As Variant = 1, Optional ByVal ovf_limit As Long = 500) As Variant
On Error GoTo Range_clean_up

    Dim ovf_buffer As Long, elem As Variant, repit As Long
    Dim col As Collection
    Dim threshold_breached As Boolean
    Dim altitude As Integer '0 = flat, -1 = descending, 1 = ascending
    Dim rtn As Variant, num As Variant
    Dim repeats As Long
    Dim factor As Long, compare_factor As Long
    Dim maxval As Variant, minval As Variant

    Set col = New Collection
    rtn = Array()
    
    If (Not IsNumeric(start_value)) Or (Not IsNumeric(end_value)) Or (Not IsNumeric(step)) Then GoTo Range_clean_up
    altitude = Switch(step = 0, 0, _
                            end_value > start_value, 1, _
                            end_value < start_value, -1, _
                            end_value = start_value, 0) '0 = flat, -1 = descending, 1 = ascending
                            
    If CBool(altitude) Then
        factor = 1
        For Each num In Array(start_value, end_value, step)
            If InStr(1, CStr(num), ".") > 0 Then
                compare_factor = 10 ^ (Len(CStr(Abs(num))) - InStr(1, CStr(Abs(num)), "."))
                If compare_factor > factor Then factor = compare_factor
            End If
        Next num
        repeats = (Abs(((end_value * factor) - (start_value * factor)) / (step * factor))) + 1
        
        If ((end_value > start_value) And step < 0) _
            Or ((end_value < start_value) And step > 0) _
            Or (step = 0) Then
                Debug.Print "Warning! Supplied parameters will not converge. Loop pre-emptively limited."
        End If
      Else
        repeats = 1
    End If
    
    ovf_buffer = 0
    elem = start_value
    Do:
        col.Add elem
        elem = elem + step

        If VarType(elem) = vbDouble Then 'Fixes fuzzy floating point math
            elem = CDbl(CStr(elem))
        End If
        threshold_breached = Choose(altitude + 2, _
                                                        (elem < end_value), _
                                                        (elem = end_value), _
                                                        (elem > end_value)) '+2 because choose is index base 1
        ovf_buffer = ovf_buffer + 1
    Loop While (Not threshold_breached) And (ovf_buffer < ovf_limit) '(elem <= end_value) And (ovf_buffer < ovf_limit)
    
    If ovf_buffer > ovf_limit Then
        Debug.Print "Warning! Loop pre-emptively stopped at ovf_limit [" & ovf_limit & "]"
    End If
    
    
    With DS
        maxval = .Maximum(Array(start_value, end_value))
        minval = .Minimum(Array(start_value, end_value))
        
        rtn = .Convert(col, "Variant()")
        rtn = .Filter(rtn, "<=", maxval)
        rtn = .Filter(rtn, ">=", minval)
    End With
Range_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Class [" & TypeName(Me) & "] method [Range]"
        Debug.Print Tab(10); "Error [" & Err.number & "]: " & Err.description
        Err.Clear
    End If

    Range = rtn
End Function


Public Function Reverse(DataStructure As Variant, Optional ByVal propagate As Boolean = False) As Variant
    Dim typeNm As String
    Dim rtn As Variant, DataStructure_arr As Variant, col As Collection
    Dim lb As Integer, ub As Integer, i As Integer
    
    If Not isDataStructure(DataStructure) Then
        If IsObject(DataStructure) Then
            Set Reverse = DataStructure
          Else
            Reverse = DataStructure
        End If
        Exit Function
    End If
    
'Initialize return container to appropriate type
    typeNm = TypeName(DataStructure)
    Select Case typeNm
        Case "Dictionary"
            Set rtn = New Scripting.Dictionary
        Case "Collection"
            Set rtn = New Collection
        Case Else
            ReDim rtn(UBound(DataStructure))
    End Select
    
'Convert input into array for easy indexing
    Select Case typeNm
        Case "Dictionary"
            DataStructure_arr = Zip(DataStructure) 'Creates nested array containing key-item pairs
        Case "Collection"
            Set col = DataStructure
            DataStructure_arr = CollectionToArray(col)
        Case Else
            DataStructure_arr = DataStructure 'redundant, but uniform
    End Select
    
'Reverse elements
    ub = UBound(DataStructure_arr)
    lb = LBound(DataStructure_arr)
    For i = lb To ub
        Select Case typeNm
            Case "Dictionary"
                If propagate And (isDataStructure(DataStructure_arr(ub - i)(1))) Then
                    rtn.Add Key:=DataStructure_arr(ub - i)(0), ITEM:=Reverse(DataStructure_arr(ub - i)(1), propagate:=True)
                  Else
                    rtn.Add Key:=DataStructure_arr(ub - i)(0), ITEM:=DataStructure_arr(ub - i)(1)
                End If
            Case "Collection"
                If propagate And (isDataStructure(DataStructure_arr(ub - i))) Then
                    rtn.Add ITEM:=Reverse(DataStructure_arr(ub - i))
                  Else
                    rtn.Add ITEM:=DataStructure_arr(ub - i)
                End If
            Case Else
                If propagate And (isDataStructure(DataStructure_arr(ub - i))) Then
                    rtn(i) = Reverse(DataStructure_arr(ub - i))
                  Else
                    rtn(i) = DataStructure_arr(ub - i)
                End If
        End Select
    Next i
    
    If IsObject(rtn) Then
        Set Reverse = rtn
      Else
        Reverse = rtn
    End If
End Function


Public Function Exists(needle As Variant, haystack As Variant, Optional ByVal wildcard_needle As Boolean = False, Optional ByVal wildcard_haystack As Boolean = False) As Boolean
    Dim match_found As Boolean, typeNm As String
    Dim search_arr As Variant
    Dim col As Collection
    Dim itm As Variant
On Error GoTo Exists_clean_up
    match_found = False
    If Not isDataStructure(haystack) Then Exists = match_found: Exit Function
    
'Standardize various possible input types for variable haystacks - normalize as array
    typeNm = TypeName(haystack)
    Select Case typeNm
        Case "Dictionary" 'Dictionary has built in method for exists, but this function can expand the method to include wildcards on either the find/needle side or the search side
            search_arr = haystack.Items
        Case "Collection"
            Set col = haystack 'reference assignment is necessary here because byref parameter can't handle a reference to a 2nd reference, to a 3rd reference, etc...
            search_arr = CollectionToArray(col)
        Case Else 'Array
            search_arr = haystack
    End Select
    
    For Each itm In search_arr
        If wildcard_needle And (Not IsObject(itm)) And (Not IsObject(needle)) Then
            If CStr(itm) Like ("*" & CStr(needle) & "*") Then
                match_found = True
                Exit For
            End If
        End If
        If wildcard_haystack And (Not IsObject(itm)) And (Not IsObject(needle)) Then
            If CStr(needle) Like ("*" & CStr(itm) & "*") Then
                match_found = True
                Exit For
            End If
        End If
        If IsObject(itm) And IsObject(needle) Then
            If itm Is needle Then
                match_found = True
                Exit For
            End If
        End If
        If (Not IsObject(itm)) And (Not IsObject(needle)) Then
            If itm = needle Then
                match_found = True
                Exit For
            End If
        End If
        If ((IsObject(itm)) And (Not IsObject(needle))) _
            Or ((Not IsObject(itm)) And (IsObject(needle))) Then
            match_found = False 'explicit, noting here that the search DataStructure might be composed of BOTH objects and non-objects,
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
    Dim val As Variant, i As Long '6/22/22..ian Integer -> Long
    Dim tmp_col As Collection
    Dim rtn As Variant

    Set tmp_col = New Collection

    If IsMissing(starting_idx) Then starting_idx = 0
    i = starting_idx
    typed_name = TypeName(enumerable)

    If Not isDataStructure(typed_name, True) Then GoTo enumerate_clean_up
    
    Select Case typed_name
        Case "Variant()", "String()", "Integer()", "Double()", _
                "Single()", "Byte()", "Long()", "Collection":
            For Each val In enumerable
                i = i + increment '6/22/22..ian Fixed increment of 1 to variable increment
                tmp_col.Add ITEM:=Array(i, val)
            Next val
        Case "Dictionary":
            For Each val In enumerable.Items
                i = i + increment '6/22/22..ian Fixed increment of 1 to variable increment
                tmp_col.Add ITEM:=Array(i, val)
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


Public Function Equivalent(DataStructure1 As Variant, DataStructure2 As Variant) As Boolean
    Dim rtn As Boolean, typeNm As Variant, i As Variant
    Dim thing As Variant, cnt As Variant, starting_idx As Variant, ending_idx As Variant
On Error GoTo Equivalent_clean_up
    Dim DataStructures As Variant
    DataStructures = Array(DataStructure1, DataStructure2)
    
    For Each thing In DataStructures
        If Not isDataStructure(thing) Then
            rtn = False
            GoTo Equivalent_clean_up
        End If
    Next thing
    
'Check if all things are DataStructures
    If DataStructureCount(DataStructure1) <> DataStructureCount(DataStructure2) Then
        rtn = False
        GoTo Equivalent_clean_up
      Else
        cnt = DataStructureCount(DataStructure1)
    End If

'Check if same types
    If TypeName(DataStructure1) <> TypeName(DataStructure2) Then
        rtn = False
        GoTo Equivalent_clean_up
      Else
        typeNm = TypeName(DataStructure1)
    End If
    
'Check for the same elements
    Select Case typeNm
        Case "Dictionary"
            If Not Equivalent(DataStructure1.keys, DataStructure2.keys) Then
                rtn = False
                GoTo Equivalent_clean_up
            End If
            
            If Not Equivalent(DataStructure1.Items, DataStructure2.Items) Then
                rtn = False
                GoTo Equivalent_clean_up
            End If
        Case Else
            
            If TypeName(DataStructure1) = "Collection" Then
                starting_idx = 1
                ending_idx = cnt
              Else
                starting_idx = 0
                ending_idx = cnt - 1
            End If
            
            For i = starting_idx To ending_idx
                If IsObject(DataStructure1(i)) Then
                    If Not (DataStructure1(i) Is DataStructure2(i)) Then
                        rtn = False
                        GoTo Equivalent_clean_up
                    End If
                  Else
                    If Not (DataStructure1(i) = DataStructure2(i)) Then
                        rtn = False
                        GoTo Equivalent_clean_up
                    End If
                End If
            Next i
    End Select
    rtn = True

Equivalent_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Err.Clear
        rtn = False
    End If
    
    Equivalent = rtn
End Function

Public Function fill(ByRef container As Variant, ByRef stuff As Variant, _
                            Optional ByVal extra_serving_size As Long = -1) As Variant
    Dim rtn As Variant
    Dim keyv As Variant
    Dim i As Integer
    Dim init_dim As Variant, add_dim As Variant
    Dim container_ As Variant
    
On Error GoTo Fill_clean_up
    If Not isDataStructure(container, accept_typename:=False) Then
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
      Else 'DataStructure
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
                
                For Each keyv In container_.keys
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
    
    If isDataStructure(container) Then
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


Public Function Filter(DataStructure As Variant, ByVal operator_expression As String, compare_against As Variant, _
                                Optional ByRef args As Variant, _
                                Optional ByVal FilterMode As eDataStructureFilterMode = eFilterTrap) As Variant

'EXAMPLE - Simple, 6/22/22
    'Given an array containing names of fruit, filter for fruit that start with the letter 'a':
'
'fruit_array = Array("Orange", "Apricot", "Apple", "Pear", "Strawberry", "Watermelon")
'fruit_starting_with_a = DS.Filter(fruit_array, "like", "A*")
    'RESULT
'>> ("Apricot", "Apple")

'EXAMPLE - Compound, 6/22/22
    'Given ALL series present on a customer's order, filter the series that belong to the sliding door department
'
'quote_series = Array("1000", "3200-T", "3070", "3070-T", "Edge_s", "3900-T", "Edge_f", "4800-T")
'filter_series = Array("10##*", "30##*", "4070-T", "Edge_s")
'
'output = DS.Filter(quote_series, "like[or]", filter_series)
    'RESULT
'>> ("1000", "3070", "3070-T", "Edge_s")

On Error GoTo Filter_clean_up
    

    Dim COLander As Collection
    Dim arr As Variant, val As Variant, thing As Variant, compared_part As Variant, pair As Variant
    Dim rtn As Variant
    Dim target_values As String, base_comparison As Variant
    Dim sub_operator As String
    Dim keep_thing As Boolean, operand_isDataStructure As Boolean
    '6/22/22..ian Added compounding operators ( <operator>[or], <operator>[and] )
    '6/22/22..ian Missing/invalid operators "out"
    '6/22/22..ian Removed unused comments
    '6/22/22..ian Added Examples (detailed documentation can still be found on the U-Drive, see top of module)

    If Not isDataStructure(DataStructure, accept_typename:=False) Then
        GoTo Filter_clean_up
      Else
        If TypeName(DataStructure) = "Dictionary" Then
            arr = DS.Zip(DataStructure)
            target_values = DS.Match(args, "=", _
                                                    "key", "key", _
                                                    "item", "item", _
                                                    args, "item")
          ElseIf TypeName(DataStructure) = "Collection" Then

            arr = Me.Convert(DataStructure, "Variant()")
            target_values = "standard"
          Else
            arr = DataStructure
            target_values = "standard"
        End If
    End If
    
    operand_isDataStructure = isDataStructure(compare_against)
    
    operator_expression = Trim(operator_expression)
    If Me.Exists(LCase(operator_expression), Array("in", "is in", "not in", "is not in", "inside", "between", "out", "outside", "beyond")) Then
        If operand_isDataStructure Then
            operand_isDataStructure = False 'Though technically a DataStructure operator, the intention is not to iterate on each individual element, but to check the operand as a whole
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
            
            If operand_isDataStructure Then
                'Check against many
                If FilterMode = eFilterTrap Then
                    'Keep going
                  Else
                    'Reverse all operands
                    'TO DO!
                End If
              Else
                If FilterMode = eFilterTrap Then
                    'Keep going
                  ElseIf FilterMode = eFilterOut Then
                    operator_expression = Match(LCase(operator_expression), "=", _
                                                                                                        "=", "<>", _
                                                                                                        "<>", "=", _
                                                                                                        "is", "is not", _
                                                                                                        "is not", "is", _
                                                                                                        ">", "<=", _
                                                                                                        "<=", ">", _
                                                                                                        "<", ">=", _
                                                                                                        ">=", "<", _
                                                                                                        "like", "not like", _
                                                                                                        "not like", "like", _
                                                                                                        "in", "not in", _
                                                                                                        "not in", "in")
                    'Reverse
                End If
                
                For Each thing In arr
                    Select Case target_values
                        Case "standard"
                            If IsObject(thing) Then Set base_comparison = thing Else base_comparison = thing
                        Case "key"
                            base_comparison = thing(0) 'key can't be an object
                        Case "item"
                            If IsObject(thing(1)) Then Set base_comparison = thing(1) Else base_comparison = thing(1)
                        Case Else: 'uh oh
                    End Select
                    
                    If Resolve(base_comparison, LCase(operator_expression), compare_against) Then
                        COLander.Add thing
                      Else
                        'Don't add
                    End If
                    
                Next thing
                
            End If

        Case ">[and]", ">=[and]", _
                "<[and]", "<=[and]", _
                "like[and]", "not like[and]", _
                "in[and]", "not in[and]", _
                "is[and]", "is not[and]", _
                "=[and]", "<>[and]" '6/22/22..ian Added compound operators
            sub_operator = Replace(operator_expression, find:="[and]", Replace:="")
            If FilterMode = eFilterOut Then
                sub_operator = DS.Match(sub_operator, "=", _
                                                            "=", "<>", "<>", "=", _
                                                            ">", "<=", ">=", "<", _
                                                            "like", "not like", _
                                                            "not like", "like", _
                                                            "in", "not in", "not in", "in", _
                                                            "is", "is not", "is not", "is")
            End If

            For Each thing In arr

                Select Case target_values
                    Case "standard": If IsObject(thing) Then Set base_comparison = thing Else base_comparison = thing
                    Case "key": base_comparison = thing(0)
                    Case "item": If IsObject(thing(1)) Then Set base_comparison = thing Else base_comparison = thing(1)
                    Case Else: 'uh oh
                End Select
                
                keep_thing = True
                For Each compared_part In compare_against
                    If Not Resolve(base_comparison, sub_operator, compared_part) Then
                        keep_thing = False
                        Exit For
                    End If
                Next compared_part
                If keep_thing Then COLander.Add thing
            Next thing
            
        Case ">[or]", ">=[or]", _
                "<[or]", "<=[or]", _
                "like[or]", "not like[or]", _
                "in[or]", "not in[or]", _
                "is[or]", "is not[or]", _
                "=[or]", "<>[or]" '6/22/22..ian Added compound operators
            sub_operator = Replace(operator_expression, find:="[or]", Replace:="")
            If FilterMode = eFilterOut Then
                sub_operator = DS.Match(sub_operator, "=", _
                                                            "=", "<>", "<>", "=", _
                                                            ">", "<=", ">=", "<", _
                                                            "like", "not like", _
                                                            "not like", "like", _
                                                            "in", "not in", "not in", "in", _
                                                            "is", "is not", "is not", "is")
            End If
            
            For Each thing In arr
            
                Select Case target_values
                    Case "standard": If IsObject(thing) Then Set base_comparison = thing Else base_comparison = thing
                    Case "key": base_comparison = thing(0)
                    Case "item": If IsObject(thing(1)) Then Set base_comparison = thing Else base_comparison = thing(1)
                    Case Else: 'uh oh
                End Select
                
                For Each compared_part In compare_against
                    If Resolve(base_comparison, sub_operator, compared_part) Then
                        COLander.Add thing
                        Exit For
                    End If
                Next compared_part
            Next thing
            
        Case Else
    End Select
    
    If IsObject(DataStructure) Then
        Set rtn = Convert(COLander, TypeName(DataStructure))
      Else
        rtn = Convert(COLander, TypeName(DataStructure))
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


Public Function Flatten(DataStructure As Variant) As Variant 'Returns a flattened array
'Note that keys belonging to collection objects cannot be preserved
    Dim rtn As Variant
    Dim nested As Variant, element As Variant, flatten_queue As Variant
    Dim flattened_in_progress As Collection
    
    If Not isDataStructure(DataStructure) Then
        rtn = Array()
        GoTo Flatten_clean_up
    End If
    
    Set flattened_in_progress = New Collection
    
    Select Case TypeName(DataStructure)
        Case "Dictionary"
            flatten_queue = DS.Zip(DataStructure)
        Case "Collection"
            flatten_queue = DS.Convert(DataStructure, "Variant()")
        Case "Variant()"
            flatten_queue = DataStructure
        Case Else
            If IsArray(DataStructure) Then 'Integer(), String(), Long(), etc... are also possible
                flatten_queue = DataStructure
              Else
                flatten_queue = Array(DataStructure) 'Just append the single element
            End If
    End Select
    
    For Each element In flatten_queue
        If isDataStructure(element) Then
            For Each nested In DS.Flatten(element)
                flattened_in_progress.Add nested
            Next nested
          Else
            flattened_in_progress.Add element
        End If
    Next element
    
    rtn = DS.Convert(flattened_in_progress, "Variant()")
    
Flatten_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Debug.Print "Class [" & TypeName(Me) & "] - Method [" & "Flatten" & "]"
        Err.Clear
        rtn = Array()
    End If

    Flatten = rtn
End Function


Public Function Formatted(arr As Variant, ByVal formatted_text As String, placeholder As Variant, _
                                        Optional ByVal substitute_null_or_invalid_with As Variant = "", _
                                        Optional ByVal vba_format As String = "", _
                                        Optional ByVal vba_firstdayofweek As Integer = 1, _
                                        Optional ByVal vba_firstweekofyear As Integer = 1) As Variant
'6/22/22..ian Release
    
    'if provided, vb_format applies the built-in VBA format function to the datastructure element BEFORE applying the [formatted_text] argument
    'vba_format represents the [Format] argument in the function syntax for Format(Expression, [Format], ...

'-----------------------------------------------------------
'-------------------    EXAMPLES ------------------------
'SIMPLE SYNTAX:
'Dim arr As Variant
'arr = Array(1, "abc", Null, -0.5, myObject)
'arr = DS.Formatted(arr, "My string is: [$]", "$")
    
'Debug.Print Join(arr, VbCrLf)
    
'>> My string is: [1]
'>> My string is: [abc]
'>> My string is: []
'>> My string is: [-0.5]
'>> My string is: []


'LAYERED SYNTAX:
'Dim first_halves As Variant, second_halves As Variant, pairs As Variant, output As Variant
'
'first_halves = Array("Jack", "Rosemary", "Bread", "Half")
'second_halves = Array("Jill", "Thyme", "Butter", "Half")
'
'pairs = DS.Zip(first_halves, second_halves)
'
'output = DS.Formatted(pairs, "$1 and $2", Array("$1", "$2"))
'Debug.Print Join(output, vbCrLf)
'>> Jack and Jill
'>> Rosemary and Thyme
'>> Bread and Butter
'>> Half and Half

'-----------------------------------------------------------

Dim copied_arr As Variant
Dim transposed As Variant
Dim imposter As Variant
Dim fill_arr As Variant, rtn As Variant
Dim i As Long, j As Long, err_count As Integer

On Error GoTo Formatted_clean_up
    
    If Not IsArray(arr) And (Not IsArray(placeholder)) Then
        imposter = arr
        rtn = Replace(formatted_text, placeholder, imposter)
        GoTo Formatted_clean_up
    End If
    
    fill_arr = DS.fill(UBound(arr) + 1, formatted_text)
    If Not IsArray(placeholder) Then
        fill_arr = apply_formats(fill_arr, placeholder, arr)
      Else
        transposed = DS.Transpose(arr)
        For i = 0 To UBound(placeholder)
            fill_arr = apply_formats(fill_arr, placeholder(i), transposed(i))
        Next i
        
    End If
    
    rtn = fill_arr
    
Formatted_clean_up:
    If Err.number <> 0 Then
        err_count = err_count + 1
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Err.Clear
        
        If err_count > 500 Then
            Exit Function
          Else
            Resume Next
        End If
    End If
    
    Formatted = rtn
End Function


'********************************************
'   IN PROGRESS
'
'Public Function Intersection(ParamArray DataStructures() As Variant) As Variant
'    'Get common elements between 2 or more DataStructure items
'    'Dictionaries & collections will be compared by items and not by keys
'    'Caution - potentially expensive operation
'    'Requires homogeneous parameters, i.e., same type
'
'    Dim master_collection As Collection
'    Dim dict As Scripting.Dictionary
'    Dim DataStructure As Variant, thing As Variant
'    Dim i As Long, j As Long
'    Dim arr As Variant, structure As Variant, main_structure As Variant
'    Dim homCol As Collection
'    Dim rtn As Variant
'
'On Error GoTo Intersection_clean_up
'    Set homCol = New Collection
'    For Each thing In DataStructures
'        homCol.Add thing
'    Next thing
'
'    Set master_collection = New Collection 'compile master_collection as a group of like-formatted data structures - copied!! treats as immutable
'    If Homogeneous(homCol) = "False" Then
'        GoTo Intersection_clean_up
'      Else
'        For Each structure In homCol
'            master_collection.Add Me.Copy(structure) 'Note the copy call here - this allows the object reference to remain unmutated
'        Next structure
'    End If
'
'    If master_collection.count <= 1 Then 'Haven't decided what to do about calling an intersection on a singular item
'        'if preventing user from calling method on a singular item, should change the parameters to DataStructure1, DataStructure2, paramarray others() as variant
'        GoTo Intersection_clean_up
'    End If
'    For i = 1 To master_collection.count
'        Set main_structure = master_collection(i)
'        For Each thing In main_structure
'            'if thing exists in each one (use the resolve method) then add it to the yes pile, otherwise add it to the no pile
'
'            'then remove the thing from all structures in master_collection
'            For j = (i + 1) To master_collection.count
'
'            Next j
'        Next thing
'    Next i
'
'    'ip
''    Select Case Homogeneous(homCol)
''        Case "False"
''            GoTo Intersection_clean_up
''        Case "Dictionary", "Collection"
''            'Proper format already
''            'COPY EACH DataStructure
''            For Each structure In homCol
''                Set thing = Me.Copy(structure)
''                master_collection.Add thing
''            Next structure
''
''        Case Else
''            For Each thing In homCol
''                master_collection.Add Me.Copy(thing)
''            Next thing
''
''    End Select
'
''Once all data structures are compiled, go through each one; sorting them into two separate piles
''Pile 1 contains items that exist in ALL of the data structures; Pile 2 contains items that don't exist in all data structures
''The main master_collection group should be un-stacked with each pass through in order to reduce the total iterations
'
'
'Intersection_clean_up:
'    If Err.number <> 0 Then
'        Debug.Print "Error [" & Err.number & "]: " & Err.description
'        Err.Clear
'    End If
'
'    If IsObject(rtn) Then
'        Set Intersection = rtn
'      Else
'        Intersection = rtn
'    End If
'    'Add to a collection of dictionaries so that the dictionaries can be paired down
'End Function
'               ^IN PROGRESS^
'********************************************************

Public Function Homogeneous(DataStructure As Variant, ParamArray others() As Variant) As Variant
'Check if all elements of the provided arguments are same-typed
'Returns the typename if uniform, False otherwise
    Dim rtn As Variant
    Dim val As Variant
    Dim tn As String
    Dim types As Scripting.Dictionary, curtype As Variant
    Dim all_args As Collection, thing As Variant
    
On Error GoTo Homogeneous_clean_up
    Set all_args = New Collection
    
    If UBound(others) >= 0 Then
        
        all_args.Add DataStructure
        For Each thing In others
            all_args.Add thing
        Next thing
    
      ElseIf isDataStructure(DataStructure) Then
        For Each thing In DataStructure
            all_args.Add thing
        Next thing
      ElseIf Not IsObject(DataStructure) Then
        all_args.Add DataStructure
      Else
        rtn = False
        GoTo Homogeneous_clean_up
    End If
    
    rtn = True 'uniform until proven otherwise
    Set types = New Scripting.Dictionary
    For Each thing In all_args
        curtype = TypeName(thing)
        If Not types.Exists(curtype) Then
            types.Add Key:=curtype, ITEM:=curtype
        End If
        
        If types.count > 1 Then
            rtn = False
            Exit For
          Else
            rtn = curtype
        End If
    Next thing

Homogeneous_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        rtn = False
        Err.Clear
    End If
    
    Set types = Nothing
    Set all_args = Nothing
    
    Homogeneous = rtn
End Function

Public Function Map(ByRef source_values As Variant, ParamArray variables() As Variant) As Boolean
On Error GoTo map_clean_up

Dim rtn As Boolean
Dim ref_var As Variant, i As Integer

If Not IsArray(source_values) Then
    For i = 0 To UBound(variables)
        On Error Resume Next
        If IsObject(source_values) Then '6/22/22..ian Map single values
            Set variables(i) = source_values
          Else
            variables(i) = source_values
        End If
        On Error GoTo map_clean_up
    Next i
    rtn = True
  Else
'    If UBound(source_values) <> UBound(variables) Then Err.Raise 9 '6/22/22..ian Remove error trap - opt for taking the min ubound
    For i = 0 To DS.Minimum(Array(UBound(source_values), UBound(variables)))  '6/22/22..ian Changed to min ubound instead of requiring equally sized arrays
        variables(i) = source_values(i)
    Next i
    rtn = True
End If

map_clean_up:
If Err.number <> 0 Then
    Debug.Print "Error in mapping multiple values across array."
    Debug.Print "Error [" & Err.number & "]: " & Err.description
End If

Map = rtn
End Function


Public Function Match(things As Variant, ByVal operator As String, _
                                    matched_value1 As Variant, output1 As Variant, _
                                    matched_value2 As Variant, output2 As Variant, _
                                    ParamArray matching_pairs() As Variant) As Variant
'See Filter function for accepted operators
On Error GoTo Match_clean_up
    Dim match_map As Scripting.Dictionary
    Dim col As Collection
    Dim match_found As Boolean, pair As Variant, thing As Variant
    Dim i As Integer
    Set match_map = New Scripting.Dictionary
    
    
    match_map.Add Key:=matched_value1, ITEM:=output1
    If Not match_map.Exists(matched_value2) Then
        match_map.Add Key:=matched_value2, ITEM:=output2
    End If
    
    If UBound(matching_pairs) > 0 Then
        i = 0
        Do While i < UBound(matching_pairs)
            On Error Resume Next
            If Not match_map.Exists(matching_pairs(i)) Then
                match_map.Add Key:=matching_pairs(i), ITEM:=matching_pairs(i + 1)
            End If
            On Error GoTo Match_clean_up
            i = i + 2
        Loop
    End If
    Set col = New Collection
    
    If isDataStructure(things, accept_typename:=False) Then

        For Each thing In things
            match_found = False
            For Each pair In Zip(match_map)
                If Resolve(thing, operator, pair(0)) Then 'See Filter function for accepted operators
                    match_found = True
                    col.Add pair(1)
                End If
                If match_found Then Exit For
            Next pair
            If Not match_found Then col.Add Empty
        Next thing
        
      Else
        match_found = False
        For Each pair In Zip(match_map)
            If Resolve(things, operator, pair(0)) Then 'See Filter function for accepted operators
                match_found = True
                col.Add pair(1)
            End If
            If match_found Then Exit For
        Next pair
    End If
    
Match_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Debug.Print "Class [" & TypeName(Me) & "] - Method [Match]"
        Err.Clear
    End If
On Error Resume Next
    If col.count = 1 Then
        If IsObject(col(1)) Then
            Set Match = col(1)
          Else
            Match = col(1)
        End If
      ElseIf col.count > 1 Then
        Select Case TypeName(things)
            Case "Collection"
                Set Match = col
            Case "Dictionary"
                'uh oh - not written yet
            Case "Variant()"
                Match = Convert(col, "Variant()")
            Case Else
                Match = Convert(col, TypeName(things))
        End Select
    End If
End Function


Public Function Merge(ParamArray DataStructures() As Variant) As Variant
    Dim merge_dictionary As Scripting.Dictionary
    Dim submerge As Variant, submerge_arr As Variant, comp As Variant
    Dim copy_DataStructures As Variant
    Dim output_type As String
    Dim i As Integer, j As Integer
    Dim basekey As String, key_ As String, addkey As String
    Dim dct As Variant, part As Variant
    
    Set merge_dictionary = New Scripting.Dictionary
    
    output_type = TypeName(DataStructures(0))
    If Not isDataStructure(output_type, accept_typename:=True) Then Exit Function
    
    For Each submerge In DataStructures
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

Public Function Minimum(DataStructure As Variant, Optional ByVal string_support As Boolean = False) As Variant
'String support will enable string length comparisons string lengths of provided elements
    Dim val As Variant
    Dim col As Collection
    Dim numeric_composition As Boolean
    Dim string_composition As Boolean
    Dim lowest As Variant
    Dim DataStructure_copy As Variant
    
    If Not isDataStructure(DataStructure) Then GoTo Minimum_clean_up
    If DataStructureCount(DataStructure) < 1 Then GoTo Minimum_clean_up
    
    Select Case TypeName(DataStructure)
        Case "Dictionary"
            DataStructure_copy = DataStructure.Items
        Case "Collection"
            Set col = DataStructure
            DataStructure_copy = CollectionToArray(col)
        Case Else
            DataStructure_copy = DataStructure
    End Select
    
    numeric_composition = True
    For Each val In DataStructure_copy
        If Not IsNumeric(val) Then
            numeric_composition = False
            Exit For
        End If
    Next val
    
    lowest = DataStructure_copy(0)
    If numeric_composition Then
        For Each val In DataStructure_copy
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

Public Function Maximum(DataStructure As Variant) As Variant
'String support will enable string length comparisons string lengths of provided elements
    Dim val As Variant
    Dim col As Collection
    Dim numeric_composition As Boolean
    Dim string_composition As Boolean
    Dim highest As Variant
    Dim DataStructure_copy As Variant
    
    If Not isDataStructure(DataStructure) Then GoTo Maximum_clean_up
    If DataStructureCount(DataStructure) < 1 Then GoTo Maximum_clean_up
    
    Select Case TypeName(DataStructure)
        Case "Dictionary"
            DataStructure_copy = DataStructure.Items
        Case "Collection"
            Set col = DataStructure
            DataStructure_copy = CollectionToArray(col)
        Case Else
            DataStructure_copy = DataStructure
    End Select
    
    numeric_composition = True
    For Each val In DataStructure_copy
        If Not IsNumeric(val) Then
            numeric_composition = False
            Exit For
        End If
    Next val
    
    highest = DataStructure_copy(0)
    If numeric_composition Then
        For Each val In DataStructure_copy
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


Public Function Outersection(ParamArray DataStructures() As Variant) As Variant
    'Returns the uncommon elements
End Function


Public Function Pop(ByRef DataStructure As Variant, Optional ByVal Index As Long) As Variant
    'Return the item indicated by the identifier; mutate the reference to exclude the element denoted by identifier
    Dim rtn As Variant, mutated As Variant
    Dim arr As Variant, comp As Variant
    Dim i As Integer, j As Integer, keyrem As Variant
    
    If Not isDataStructure(DataStructure) Then
        GoTo Pop_clean_up
    End If
    
    Select Case TypeName(DataStructure)

        Case "Dictionary"
            If IsMissing(Index) Then Index = 0
            If Index = -1 Then Index = DataStructure.count - 1 'Last element
            
            If DataStructure.count >= Index + 1 Then
                arr = Me.Zip(DataStructure.keys, DataStructure.Items)
                keyrem = arr(Index)(0)
                If IsObject(arr(Index)(1)) Then
                    Set rtn = arr(Index)(1)
                  Else
                    rtn = arr(Index)(1)
                End If
                DataStructure.Remove Key:=keyrem
            End If
            
        Case "Collection"
            If IsMissing(Index) Then Index = 1
            If Index = -1 Then Index = DataStructure.count
            If Index = 0 Then Index = 1 'Collections are base 1
            If DataStructure.count >= Index Then
                If IsObject(DataStructure(Index)) Then
                    Set rtn = DataStructure(Index)
                  Else
                    rtn = DataStructure(Index)
                End If
            End If
            
            If DataStructure.count > 0 Then
                Set DataStructure = Me.Remove(DataStructure, Index - 1) 'Method [Remove] uses index base 0
            End If

        Case Else 'Array
            If IsMissing(Index) Then Index = 0
            If Index = -1 Then Index = UBound(DataStructure)
            If UBound(DataStructure) = 0 Then
                rtn = DataStructure(0)
                DataStructure = Array()
                GoTo Pop_clean_up
            End If
            ReDim mutated(UBound(DataStructure) - 1)
            rtn = DataStructure(Index)
            j = 0 'index of last addition to mutated array, starting from 0
            For i = 0 To UBound(DataStructure)
                If i <> Index Then
                    mutated(j) = DataStructure(i)
                    j = j + 1
                End If
            Next i
            If j <> i - 1 Then
                'uh oh
                'don't overwrite
              Else
                DataStructure = mutated
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


Public Function Postfixed(postfixables As Variant, ByVal postfix As String, Optional ByVal apply_to_dict_keys As Boolean = False, Optional ByVal subst_placeholder As String = "") As Variant
    If IsObject(postfixables) Then
        Set Postfixed = outerfix(postfixables, _
                                            attachment:=postfix, _
                                            prefix:=False, _
                                            apply_to_dict_keys:=apply_to_dict_keys, _
                                            substitute:=subst_placeholder)
      Else
        Postfixed = outerfix(postfixables, _
                                    attachment:=postfix, _
                                    prefix:=False, _
                                    apply_to_dict_keys:=apply_to_dict_keys, _
                                    substitute:=subst_placeholder)
    End If
End Function

Public Function Prefixed(prefixables As Variant, ByVal prefix As String, Optional ByVal apply_to_dict_keys As Boolean = False, Optional ByVal subst_placeholder As String = "") As Variant
    If IsObject(prefixables) Then
        Set Prefixed = outerfix(prefixables, _
                                            attachment:=prefix, _
                                            prefix:=True, _
                                            apply_to_dict_keys:=apply_to_dict_keys, _
                                            substitute:=subst_placeholder)
      Else
        Prefixed = outerfix(prefixables, _
                                    attachment:=prefix, _
                                    prefix:=True, _
                                    apply_to_dict_keys:=apply_to_dict_keys, _
                                    substitute:=subst_placeholder)
    End If
End Function


Public Function Remove(DataStructure As Variant, Optional ByVal Index As Integer = 0) As Variant
    Dim col As Collection
    Dim i As Integer
    Dim arr As Variant, rtn As Variant
    Dim rkey As String
    
    If Not isDataStructure(DataStructure) Then
        On Error Resume Next
        Remove = DataStructure
        If Err.number <> 0 Then Err.Clear: Set Remove = DataStructure
        Exit Function
    End If
    
    Select Case TypeName(DataStructure)
        Case "Dictionary"
            rkey = DataStructure.keys(Index)
            Set rtn = DataStructure.Remove(rkey)
        Case "Collection"
            Set col = DataStructure
            arr = CollectionToArray(col)
            Set rtn = ArrayToCollection(Me.Remove(arr, Index))
        Case Else
            Set col = New Collection
            
            For i = 0 To UBound(DataStructure)
                If i <> Index Then col.Add DataStructure(i)
            Next i
            rtn = CollectionToArray(col)
    End Select
    
    If IsObject(rtn) Then
        Set Remove = rtn
      Else
        Remove = rtn
    End If
End Function

Public Function Transpose(arr As Variant) As Variant
'6/22/22..ian Release

' m x n -> n x m
'i.e.,
' ( (1, 2, 3, 4), _
    (5, 6, 7, 8) ) [m=2 x n=4]
' Becomes
' ( (1, 5), _
    (2, 6), _
    (3, 7), _
    (4, 8) ) [m=4, n=2]
    
    Dim rtn As Variant, i As Long, j As Long
    Dim i_col As Collection, j_col As Collection
    Dim err_count As Long
    Dim is2D As Boolean
    
    On Error Resume Next
    is2D = IsArray(arr) And IsArray(arr(0))
    If Not is2D Or Err.number <> 0 Then
        rtn = arr
        Exit Function
    End If
    
On Error GoTo Transpose_clean_up
    Set j_col = New Collection
    
    For j = 0 To UBound(arr(0))
        Set i_col = New Collection
        
        For i = 0 To UBound(arr)
            i_col.Add arr(i)(j)
        Next i
        j_col.Add DS.Convert(i_col, "Variant()")
    Next j
    
Transpose_clean_up:
    If Err.number <> 0 Then
        err_count = err_count + 1
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Err.Clear
        If err_count > 500 Then
            'Pass - don't resume
            Exit Function
          Else
            Resume Next
        End If
    End If
    
    Transpose = DS.Convert(j_col, "Variant()")
End Function

Public Function Zip(ParamArray DataStructures() As Variant) As Variant
    On Error GoTo Zip_clean_up
    Dim rtn As Variant, copy_DataStructures As Variant, DataStructure As Variant, thing As Variant
    Dim output_col As Collection, tcol As Collection
    Dim i As Integer, j As Integer, thisCount As Integer, lastCount As Integer, itCount As Integer

'Make sure all parameters are the same count/size
    lastCount = -2 'arbitrary, but not -1 since ubound of empty array is -1
    For Each DataStructure In DataStructures
        thisCount = DataStructureCount(DataStructure)
        
        If lastCount = -2 Then
            lastCount = thisCount
        End If
        
        If thisCount <> lastCount Then
            'Incompatible
            Zip = Empty
            Exit Function
        End If
    Next DataStructure
    itCount = thisCount 'Use later to output and size the return array to the span (itCount) of the input DataStructures
    
    'Special case for standalone dictionaries; dictionaries are compacted and returned
    If UBound(DataStructures) = 0 Then
        If TypeName(DataStructures(0)) = "Dictionary" Then
            With DataStructures(0)
                rtn = zipArray_(.keys, .Items)
            End With
            Zip = rtn
            Exit Function
        End If
    End If
    
    ReDim copy_DataStructures(UBound(DataStructures))
    For i = 0 To UBound(DataStructures)
        On Error Resume Next
        copy_DataStructures(i) = DataStructures(i)
        If Err.number <> 0 Then
            Err.Clear
            If TypeName(DataStructures(i)) = "Collection" Then
                Set copy_DataStructures(i) = DataStructures(i)
              ElseIf TypeName(DataStructures(i)) = "Dictionary" Then
                copy_DataStructures(i) = Zip(DataStructures(i))
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
        For Each thing In copy_DataStructures
            Select Case TypeName(thing)
                Case "Collection"
                    tcol.Add ITEM:=thing(j + 1)
                Case "Dictionary" 'uh oh
                    tcol.Add ITEM:=thing(thing.keys(j))
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

Private Function apply_formats(expressions As Variant, ByVal placeholder As String, replacements As Variant) As Variant
    Dim copied_things As Variant
    Dim i As Long
    Dim replacement As Variant
    copied_things = DS.Copy(expressions)

    For i = 0 To UBound(copied_things)
        copied_things(i) = Replace(copied_things(i), placeholder, CStr(replacements(i)))
    Next i

    apply_formats = copied_things
End Function

Private Function DataStructureCount(ByRef DataStructure As Variant) As Integer
On Error GoTo DataStructureCount_clean_up
    Dim i As Integer, val As Variant
    i = 0
    If TypeName(DataStructure) = "Dictionary" Then
        For Each val In DataStructure.Items
            i = i + 1
        Next val
      Else
        For Each val In DataStructure
            i = i + 1
        Next val
    End If
DataStructureCount_clean_up:
    If Err.number <> 0 Then Err.Clear
'    If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "]: " & Err.description
    DataStructureCount = i
End Function


Private Function DataStructureMaxSpan(ByRef DataStructure As Variant) As Long

End Function


Private Function DataStructureToArray(DataStructure As Variant) As Integer

End Function


Private Function isDataStructure(DataStructure As Variant, Optional ByVal accept_typename As Boolean = True) As Boolean
    Dim typnm As String
    Dim rtn As Boolean: rtn = False
'Note:
'       Accepts Syntax:
'1.           isDataStructure("Variant()")
'2.          isDataStructure(object)
    If accept_typename And TypeName(DataStructure) = "String" Then
        typnm = DataStructure
      Else
        typnm = TypeName(DataStructure)
    End If
    Select Case typnm 'Accepts string name of DataStructure type as well as the DataStructure itself
        Case "Variant()", "Long()", _
                    "Single()", "Byte()", _
                    "Integer()", "Double()", _
                    "String()", "Boolean()", _
                    "Collection", "Dictionary"
            rtn = True
        Case Else
            rtn = False
    End Select
    isDataStructure = rtn
End Function


Private Function CollectionToArray(DataStructure_col As Collection) As Variant
    Dim arr As Variant
    Dim i As Integer
    If DataStructure_col.count > 0 Then
        ReDim arr(DataStructure_col.count - 1)
        For i = 0 To UBound(arr)
            If IsObject(DataStructure_col(i + 1)) Then
                Set arr(i) = DataStructure_col(i + 1)
              Else
                arr(i) = DataStructure_col(i + 1)
            End If
        Next i
      Else
        arr = Array()
    End If
    CollectionToArray = arr
End Function


Private Function DictionaryToArray(DataStructure_dictionary As Scripting.Dictionary) As Variant
    Dim comp As Variant, keyarr As Variant, itemarr As Variant
    Dim arr() As Variant
    Dim i As Integer
    If DataStructure_dictionary.count > 0 Then
        ReDim arr(DataStructure_dictionary.count - 1)
        keyarr = DataStructure_dictionary.keys
        itemarr = DataStructure_dictionary.Items
        For i = 0 To (DataStructure_dictionary.count - 1)
            arr(i) = Array(keyarr(i), itemarr(i))
        Next i
    End If
    DictionaryToArray = arr
End Function


Private Function ArrayToCollection(DataStructure As Variant, Optional ByVal populateKeys As Boolean = True, Optional keys As Variant, Optional ByVal auto_increment_dup_keys As Boolean = True) As Collection
    Dim rtn_col As Collection
    Dim val As Variant, pair As Variant, paired As Variant
    Dim ckey As Variant, citem As Variant, i As Integer, newkey As Variant
    Dim key_source As Variant
    Dim source_is_valid As Boolean, use_keys As Boolean
    Set rtn_col = New Collection
    
    If IsArray(DataStructure) Then
        On Error Resume Next
        If UBound(DataStructure) >= 0 Then
            If Err.number <> 0 Then
                GoTo ArrayToCollection_clean_up
            End If
        End If
    End If

    If populateKeys Then
        'Determine where to get the keys (either nested or provided as an additional argument)
        If IsMissing(keys) Then
            key_source = "nested"
          Else
            key_source = "optional arg"
        End If
        
        source_is_valid = True
        Select Case key_source
            Case "nested"
                'Validate the nested keys
                For Each pair In DataStructure
                    source_is_valid = (IsArray(pair))
                    If source_is_valid Then source_is_valid = (UBound(pair) = 1)
                    If source_is_valid Then source_is_valid = (Len(CStr(pair(0))) > 0)
                    If Not source_is_valid Then Exit For
                Next pair
                
            Case "optional arg"
                'Validate the provided keys array
                source_is_valid = (UBound(keys) = (UBound(DataStructure)))
                If source_is_valid Then
                    For Each ckey In keys
                        source_is_valid = Len(CStr(ckey)) > 0
                        If Not source_is_valid Then Exit For
                    Next ckey
                End If
                
            Case Else
                source_is_valid = False
                'uh oh
        End Select
        
        If source_is_valid Then
            Select Case key_source
                Case "nested"
                    paired = DataStructure
                    use_keys = True
                Case "optional arg"
                    paired = DS.Zip(keys, DataStructure)
                    use_keys = True
            End Select
          Else
            use_keys = False
        End If
      Else
        use_keys = False
    End If
                          
    If use_keys Then
        For Each pair In paired
            On Error Resume Next
            i = 0
            newkey = CStr(pair(0))
            Do
                Err.Clear
                rtn_col.Add ITEM:=pair(1), Key:=newkey
                newkey = CStr(pair(0)) & "-" & CStr(i)
                i = i + 1
            Loop While Err.number = 457
            
        Next pair
        On Error GoTo ArrayToCollection_clean_up
      Else
        
        'add items with no keys
        For Each val In DataStructure
            rtn_col.Add val
        Next val
    End If

    
ArrayToCollection_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Debug.Print "Class [" & TypeName(Me) & "] - Method [" & "ArrayToCollection" & "]"
        Err.Clear
    End If
    
    Set ArrayToCollection = rtn_col
    
End Function


Private Function ArrayToDictionary(DataStructure As Variant, Optional ByVal populateKeys As Boolean = True) As Scripting.Dictionary
    Dim rtn_dictionary As Scripting.Dictionary
    Dim nextkey As String, basekey As String
    Dim val As Variant, comp As Variant
    Dim keyidx_suffix As Integer, i As Integer
    Set rtn_dictionary = New Scripting.Dictionary
    
    If IsArray(DataStructure) Then
        If IsArray(DataStructure(0)) Then
            If populateKeys Then
                For Each comp In DataStructure
                    keyidx_suffix = 0
                    basekey = CStr(comp(0))
                    nextkey = basekey
                    While rtn_dictionary.Exists(nextkey)
                        nextkey = basekey & CStr(keyidx_suffix)
                        keyidx_suffix = keyidx_suffix + 1
                    Wend
'                    rtn_dictionary.Add Key:=nextkey, Item:=Remove(comp, Index:=0)
                    rtn_dictionary.Add Key:=nextkey, ITEM:=comp(1)
                Next comp
              Else 'ignore keys, use regular indexing
                i = 0
                For Each val In DataStructure
                    rtn_dictionary.Add Key:=i, ITEM:=val
                    i = i + 1
                Next val
            End If 'populate keys
          Else
            i = 0
            For Each val In DataStructure
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


Private Function outerfix(stringables As Variant, ByVal attachment As String, Optional ByVal prefix As Boolean = True, Optional ByVal apply_to_dict_keys As Boolean = False, Optional ByVal substitute As String = "") As Variant
    Dim col As Collection
    Dim unfixed As String, tmpFixed, i As Integer
    Dim elem As Variant, elements As Variant
    Dim output As Variant
    Dim amended_attachment As String
    
On Error GoTo outerfix_clean_up
    If TypeOf stringables Is Collection Then GoTo outerfix_clean_up 'Won't be able to return an equivalent Collection since Collection keys and items are immutable
    
    If isDataStructure(stringables) Then
        Set col = New Collection
        If TypeName(stringables) = "Dictionary" Then
            elements = IIf(apply_to_dict_keys, stringables.keys, stringables.Items)
          ElseIf TypeName(stringables) = "Collection" Then
            GoTo outerfix_clean_up
          Else
            elements = stringables
        End If
    End If

    For Each elem In elements
        If Len(substitute) > 0 Then
            amended_attachment = Replace(attachment, find:=substitute, Replace:=CStr(elem))
          Else
            amended_attachment = attachment
        End If
        unfixed = ""
        On Error Resume Next
        unfixed = CStr(elem) 'Error may occur on this line for non-convertable variables, e.g., objects
        If Err.number <> 0 Then
            Err.Clear
            tmpFixed = ""
          Else
            tmpFixed = IIf(prefix, amended_attachment & unfixed, unfixed & amended_attachment)
        End If
        On Error GoTo outerfix_clean_up
        col.Add tmpFixed
    Next elem
    
    If IsObject(stringables) Then Set output = Copy(stringables) Else output = Copy(stringables)
    
    Select Case TypeName(output)
        Case "Dictionary" 'Copy stringables so that the values in the reference are unchanged
            For i = 0 To UBound(elements)
                If col(i + 1) = "" Then
                    'Pass
                  Else
                    If apply_to_dict_keys Then
                        output.Key(elements(i)) = col(i + 1) 'Collection is index base 1
                      Else 'apply to dictionary items instead
                        output.ITEM(elements(i)) = col(i + 1)
                    End If
                End If 'Unable to convert value in code block above this one
            Next i
        Case "Collection"
            GoTo outerfix_clean_up 'Invalid - collections are immutable; returning a collection without retaining the keys may be undesirable
        Case Else
            For i = 0 To UBound(stringables)
                If col(i + 1) = "" Then
                    'Pass
                  Else
                    output(i) = col(i + 1)
                End If
            Next i
    End Select
    
outerfix_clean_up:
    If Err.number <> 0 Then
        Debug.Print "Error [" & Err.number & "]: " & Err.description
        Debug.Print Tab(10); "Class [" & TypeName(Me) & "] - Method [" & "outerfix" & "]"
        Err.Clear
        If IsObject(stringables) Then Set output = Nothing
    End If
    
    If IsObject(stringables) Then
        Set outerfix = output
      Else
        outerfix = output
    End If
End Function


Private Function Resolve(left_side As Variant, operator As String, right_side As Variant) As Boolean
On Error GoTo Resolve_clean_up
    Dim vali As Variant
    Dim lowercase_operator As String, mini_ As Variant, maxi_ As Variant
    lowercase_operator = LCase(operator)
    Dim rtn As Boolean
    
    rtn = False
    
    Select Case lowercase_operator
        Case "=" 'Keeping equal sign operator as values only since it feels wrong to say some object = another object; even if their sub-object values are the same, their references might point to different places in memory
            rtn = (left_side = right_side)
        Case "<>"
            rtn = (left_side <> right_side)
        Case "is" 'Objects only <- REDACTED - it would be more convenient to abstract away the distinction between objects and values to allow for comparisons to read like "if count is 1 then..."
            If IsObject(left_side) And IsObject(right_side) Then
                rtn = (left_side Is right_side)
              ElseIf Not IsObject(left_side) And Not IsObject(right_side) Then
                rtn = (left_side = right_side)
              Else
                rtn = False
            End If
        Case "is not"
            If IsObject(left_side) And IsObject(right_side) Then
                rtn = (left_side Is Not right_side)
              ElseIf Not IsObject(left_side) And Not IsObject(right_side) Then
                rtn = (left_side <> right_side)
              Else
                rtn = False
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
        Case "not in", "is not in" '6/22/22..ian Added missing "not in" operator
            rtn = Not DS.Exists(left_side, right_side)
        Case "like"
'            rtn = (CStr(left_side) Like ("*" & CStr(right_side) & "*"))
            rtn = (CStr(left_side) Like (CStr(right_side)))
        Case "not like"
            rtn = (Not (CStr(left_side) Like (CStr(right_side))))
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

Private Function get_opposite_operator(ByVal operator As String, Optional ByVal unmatched_default As String = "") As String
    get_opposite_operator = DS.Match(operator, "=", _
                                                            "=", "<>", "<>", "=", _
                                                            "like", "not like", "not like", "like", _
                                                            ">", "<=", "<=", ">", ">=", "<", "<", ">=", _
                                                            "is", "is not", "is not", "is", _
                                                            "in", "not in", "not in", "in", _
 _
                                                            "=[and]", "<>[or]", "<>[or]", "=[and]", _
                                                            "=[or]", "<>[and]", "<>[and]", "=[or]", _
                                                            ">[and]", "<=[or]", "<=[or]", ">[and]", _
                                                            ">=[and]", "<[or]", "<[or]", ">=[and]", _
                                                            ">[or]", "<=[and]", "<=[and]", ">[or]", _
                                                            ">=[or]", "<[and]", "<[and]", ">=[or]", _
 _
                                                            "like[and]", "not like[or]", "not like[or]", "like[and]", _
                                                            "like[or]", "not like[and]", "not like[and]", "like[or]", _
 _
                                                            "is[and]", "is not[or]", "is not[or]", "is[and]", _
                                                            "is not[and]", "is[or]", "is[or]", "is not[and]", _
 _
                                                            operator, unmatched_default)
End Function
