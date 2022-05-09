Attribute VB_Name = "Module1"
Option Explicit

Function vowel_shift(ByVal char As String) As String
    vowel_shift = IIf(InStr("aeiou", LCase(char)) > 0, Chr(Asc(char) + 1), char)

End Function

Sub Test_map()


Dim arr As Variant
Dim shifted_arr As Variant
Dim obfuscated_text As Variant
Dim quote As String
quote = "All that is gold does not glitter, " & _
        "Not all those who wander are lost"
arr = DS.CharacterArray(quote)
shifted_arr = DS.Apply(arr, "vowel_shift", 0)
obfuscated_text = Join(shifted_arr, "")
Debug.Print obfuscated_text
End Sub

Sub test_convert_from_array_with_keys()
    Dim arr As Variant
    Dim col As Collection
    Dim keys_arr As Variant, dkey As Variant
    Dim items_arr As Variant
    Dim output As Variant
    Dim dict As Scripting.Dictionary
    
    keys_arr = Array("a", "b", "c", "d")
    items_arr = Array(1, 2, 3, 4)
    
    Set dict = DS.Convert(items_arr, "Dictionary", keys:=keys_arr)
    For Each dkey In keys_arr
        Debug.Print dict(dkey)
    Next dkey
    Stop
    
    Set dict = Nothing
    
    
    
End Sub


Function force_at_yield(ByVal yield_strength As Double, ByVal area_mom_inert As Double, ByVal neutral_axis_dist As Double, ByVal length As Double) As Double
    force_at_yield = (yield_strength * area_mom_inert) / (neutral_axis_dist * length)
End Function

Function deflection_at_yield(ByVal force As Double, ByVal length As Double, ByVal elast_mod As Double, ByVal area_mom_inert As Double) As Double
    deflection_at_yield = (force * (length ^ 3)) / (3 * elast_mod * area_mom_inert)
End Function

Sub test_material_props_dictionary()
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
l1 = 1 * 12 '36 in.
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

End Sub

Sub test_boolean_and_string_arrays()
    Dim ingredients As Variant
    Dim in_stock As Variant
    Dim sample_col As Variant, sample_dict As Variant, sample_arr As Variant
    
    ingredients = Array("macaroni shells", "pesto sauce", "bell peppers", "chicken", "jalapenos", "mozarella")
    in_stock = Array(True, False, True, True, False, False)
    
    Set sample_col = DS.Convert(in_stock, "Collection", keys:=ingredients)
    Debug.Print "Variable sample_col type: " & TypeName(sample_col)
    Debug.Print "# of items in sample_col: " & sample_col.Count
    
    Set sample_dict = DS.Convert(in_stock, "Dictionary", keys:=ingredients)
    Debug.Print "Variable sample_dict type: " & TypeName(sample_dict)
    Debug.Print "# of items in sample_dict: " & sample_dict.Count
    
    sample_arr = DS.Convert(sample_dict, "Variant()")
    Debug.Print "Elements in array: " & UBound(sample_arr) + 1 'Base 0
'    debug.Print "Elements: " & Join(DS.Apply("Join"
    Debug.Print "Elements in each sub-array: " & UBound(sample_arr(0)) + 1 'Base 0
    
    Set sample_col = Nothing
    Set sample_dict = Nothing
    
End Sub
