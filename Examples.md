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
