

Public Function sqlTextLocal( _
    nm As String _
) As String
    sqlTextLocal = _
    sqlTextInProject( _
    nm, vbProjectLocal())
'Debug.Print cnGnsDoyle().Execute(sqlTextLocal("sqlOf_ERC_PTOSIZE")).GetString 'sqlOf_428_MOTORVOLTAGE
End Function

Public Function sqlOf_() As String
#If False Then
'''SQL'''
-- SQL STATEMENT

; --
'''SQL'''
#End If
    sqlOf_ = sqlTextLocal( _
        "sqlOf_" _
    )
End Function

Public Function sqlOf_simpleSelWhere( _
    FromView As String, GetField As String, _
    WhereField As String, Matches As Variant _
) As String
    Dim mtExpr As String
    
    '   generate filter expression
    '   based on type of supplied
    '   matching value
    If IsArray(Matches) Then
        mtExpr = " in ('" & Join(Matches, "', '") & "') "
        Stop
    ElseIf IsNumeric(Matches) Then
        mtExpr = " = " & CStr(Matches) & " "
    ElseIf VarType(Matches) = vbString Then
        ''' for String data, single quotes
        ''' must be "escaped" by repeating
        ''' each instance, that is, replace
        ''' each one with two of the same.
        mtExpr = " = '" & Replace( _
            Matches, "'", "''" _
        ) & "' "
    ElseIf IsDate(Matches) Then
        mtExpr = " = #" & CStr(Format$(Matches, "yyyy/mm/dd")) & "# "
        Stop
    ElseIf IsNull(Matches) Then
        mtExpr = " is null "
    Else
        Stop
    End If
    '   NOTE: this block MIGHT want
    '   exported to its own function
    
    sqlOf_simpleSelWhere _
        = " select " & GetField _
        & " from " & FromView _
        & " where " & WhereField _
        & mtExpr & _
    ";"
End Function

Public Function sqlOf_gnsMatlSpec1ops() As String
    sqlOf_gnsMatlSpec1ops = sqlOf_gnsMatlSpec1ops_v0_1()
End Function

Public Function sqlOf_gnsMatlSpec1ops_v0_1() As String
#If False Then
'''SQL'''
-- SQL STATEMENT
select i.Specification1 form
from vgMfiItems i
where i.Family in ('D-BAR', 'DSHEET')
  and ISNULL(i.Specification16, '') = ''
  and i.Specification1 is not null
group by i.Specification1
order by CHARINDEX(
  LEFT(i.Specification1 + '--', 2),
  ' BA TU PI LE CH SP SH ST -- '
)
; --
'''SQL'''
#End If
    sqlOf_gnsMatlSpec1ops_v0_1 = sqlTextLocal( _
        "sqlOf_gnsMatlSpec1ops_v0_1" _
    )
End Function

Public Function sqlOf_MatlSpecXref() As String
    sqlOf_MatlSpecXref = sqlOf_MatlSpecXref_v0_1()
End Function

Public Function sqlOf_MatlSpecXref_v0_1() As String
#If False Then
'''SQL'''
-- SQL STATEMENT
with q1 as (
  select *
  from vgMfiItems i
  where i.Family in ('D-BAR', 'DSHEET')
    and ISNULL(i.Specification16, '') = ''
)
, q2 as (
select MIN(i.Item) RefItem --i.Item --, i.Family
, i.Specification1
, i.Specification2
, i.Specification3
, i.Specification4
, i.Specification5
, i.Specification6
, i.Specification7
, i.Specification8
, i.Specification9
, i.Specification10
, i.Specification11
, i.Specification12
, i.Specification13
, i.Specification14
, i.Specification15
--, i.Specification16
from q1 i
-- where i.Family in ('D-BAR', 'DSHEET')
--   and ISNULL(i.Specification16, '') = ''
group by i.Family
, i.Specification1
, i.Specification2
, i.Specification3
, i.Specification4
, i.Specification5
, i.Specification6
, i.Specification7
, i.Specification8
, i.Specification9
, i.Specification10
, i.Specification11
, i.Specification12
, i.Specification13
, i.Specification14
, i.Specification15
)
, q3 as (
select u.RefItem, u.val --, u.col
from q2 unpivot (val for col in (
  q2.Specification1
, q2.Specification2
, q2.Specification3
, q2.Specification4
, q2.Specification5
, q2.Specification6
, q2.Specification7
, q2.Specification8
, q2.Specification9
, q2.Specification10
, q2.Specification11
, q2.Specification12
, q2.Specification13
, q2.Specification14
, q2.Specification15
)) u
where u.val <> ''
)
, q4 as (
  select q3.val, a.val also
  from q3 join q3 a
    on q3.RefItem = a.RefItem
   and q3.val <> a.val
)
, qZ as (
  select *
  from q4
)
--
--
select * from qZ
group by qZ.val, qZ.also
order by qZ.val, qZ.also
--.Item
--, u.Family
--, u.Specification1 Form
--
--
; --
'''SQL'''
#End If
    sqlOf_MatlSpecXref_v0_1 = sqlTextLocal( _
        "sqlOf_MatlSpecXref_v0_1" _
    )
End Function

Public Function sqlOf_GnsPartInfo( _
    Item As String _
) As String
    sqlOf_GnsPartInfo = Replace( _
        sqlTextLocal( _
        "sqlOf_GnsPartInfo" _
        ), "%%%", Item _
    )
#If False Then
'''SQL'''
-- GnsPartInfo
with t as (
    select iType, bomStr
    from (values
        ('M', 51970), -- kNormalBOMStructure
        ('R', 51973)  -- kPurchasedBOMStructure
    ) ls(iType, bomStr)
)
select
  i.Item [Part Number]
, i.Family    [Cost Center]
, i.Type
, t.bomStr
, i.Weight     GeniusMass
, i.Width      Extent_Width
, i.Length     Extent_Length
, i.Diameter   Extent_Area
, i.Thickness
, i.Height
from vgMfiItems i
     inner Join
     t on i.Type = t.iType
where i.Item = '%%%'
; --
'''SQL'''
#End If
End Function

Public Function sqlOf_GnsPartMatl( _
    Item As String _
) As String
    sqlOf_GnsPartMatl = Replace( _
        sqlTextLocal( _
        "sqlOf_GnsPartMatl" _
        ), "%%%", Item _
    )
#If False Then
'''SQL'''
-- GnsPartMatl
select
  B.ItemOrder Ord
, b.Item RM -- was Material
, m.Family MtFamily
, b.QuantityInConversionUnit RMQTY -- was Qty
, b.ConversionUnit RMUNIT -- was Unit
from vgIcoBillOfMaterials b
     inner join vgMfiItems m
       on b.Item = m.Item
where b.Product = '%%%'
; --
'''SQL'''
#End If
End Function

Public Function sqlOf_GnsMatlOptions( _
    Matl As String, Dims As Variant _
) As String
    sqlOf_GnsMatlOptions = _
    sqlOf_GnsMatlOptions_v0_2( _
        Matl, Dims _
    )
End Function

Public Function sqlOf_GnsMatlOptions_v0_1( _
    Matl As String, Wdth As Double, Hght As Double, _
    Optional Thck As Double = -1, _
    Optional Lgth As Double = 0 _
) As String
    ''' DON'T try to do anything with this yet!
    ''' see notes on where things are
    sqlOf_GnsMatlOptions_v0_1 = _
    Replace(Replace(Replace(Replace(Replace( _
        sqlTextLocal("sqlOf_GnsMatlOptions_v0_1") _
        , "$MTL$", Matl) _
        , "#THK#", "") _
        , "#WID#", "") _
        , "#HGT#", "") _
        , "#LNG#", "") _
    '''
#If False Then
'''SQL'''
with sp as (
  select mtl, thk, wid, hgt, lng
  from (values (
    '$MTL$'
    --, 1.5, 2.5, 0.1875
    , CONVERT(float, #THK#)
    , CONVERT(float, #WID#)
    , CONVERT(float, #HGT#)
    , CONVERT(float, #LNG#)
  )) sp(mtl, thk, wid, hgt, lng)
)
, mt as (
  select i.Item
  , i.Thickness, i.Width, i.Length, i.Height
  , i.Description1
  , i.Specification1
  , i.Specification2
  , i.Specification3
  , i.Specification4
  --, i.Specification5
  , i.Specification6
  from vgMfiItems i
       inner join sp
       on i.Specification6 = sp.mtl
  where i.Family in ('D-BAR', 'DSHEET')
)
/*
*/
, md as (
  select u.Item, u.Axis, u.Dim
  , u.Dim + 0.0006 dmx
  , u.Dim - 0.0006 dmn
  from mt m
  unpivot (Dim for Axis in (
    m.Thickness , m.Width, m.length, m.Height
  )) u
  where u.Dim > 0
)
, pd as (
  select u.Axis, u.Dim
  from sp m
  unpivot (Dim for Axis in (
    m.thk, m.wid, m.lng --, m.hgt
  )) u
  where u.Dim > 0
)
, dx as (
  select md.Item, pd.Dim
  , pd.Axis px
  , md.Axis mx
  from pd inner join md
    --on pd.Dim = md.Dim
    on pd.Dim between md.dmn and md.dmx
)
--
--
select dx.Item
, COUNT(dx.Dim) MatchCt
from dx
group by dx.Item
order by MatchCt desc, dx.Item
--
--
'''SQL'''
#End If
End Function

Public Function sqlOf_GnsMatlOptions_v0_2( _
    Matl As String, Dims As Variant _
) As String
    ''' DON'T try to do anything with this yet!
    ''' see notes on where things are
    If IsArray(Dims) Then
        sqlOf_GnsMatlOptions_v0_2 = _
        Replace(Replace( _
            sqlTextLocal("sqlOf_GnsMatlOptions_v0_2") _
            , "%%S6%%", Matl) _
            , "%%LS%%", Join(Dims, "), (")) _
        '''
    ElseIf IsNumeric(Dims) Then
        sqlOf_GnsMatlOptions_v0_2 = _
        sqlOf_GnsMatlOptions_v0_2( _
            Matl, Array(Dims) _
        )
    Else
        Stop 'because this might be an issue
        ''' will resort to a sane default for now
        sqlOf_GnsMatlOptions_v0_2 = _
        sqlOf_GnsMatlOptions_v0_2( _
            Matl, Array(0.075) _
        ) 'should pick up 14GA sheet metal only
        '  might switch to something better
        '  matching structural materials
    End If
#If False Then
'''SQL'''
-- SQL STATEMENT
--
with mt as (
  select i.Item, i.Family
  , i.Thickness, i.Width, i.Length, i.Height, i.Diameter
  , i.Description1
  , i.Specification1
  , i.Specification2
  , i.Specification3
  , i.Specification4
  --, i.Specification5
  --, i.Specification6
  from vgMfiItems i
  --     inner join sp
  --     on i.Specification6 = sp.mtl
  where i.Family in ('D-BAR', 'DSHEET')
    and i.Specification6 = '%%S6%%' -- MS
)
, md as (
  select u.Item, u.Axis, u.Dim
  , u.Dim + 0.001 dmx
  , u.Dim - 0.001 dmn
  from mt m
  unpivot (Dim for Axis in (
    m.Thickness , m.Width, m.length, m.Height, m.Diameter
  )) u
  where u.Dim > 0
)
, dx as (
  select md.Item
  , md.Axis mx
  , d.v Dim
  from (values (%%LS%%)) as d(v) -- 0.25), (2), (3
  inner join md on d.v between md.dmn and md.dmx
)
, mc as (
  select dx.Item
  , COUNT(dx.Dim) MatchCt
  from dx
  group by dx.Item
)
--
--
select mc.Item, mc.MatchCt
, mt.Family
, mt.Specification1
, mt.Specification2
--, mt.Specification3
--, mt.Specification4
, mt.Description1
from mc inner join mt
on mc.Item = mt.Item
order by mt.Specification1
, mc.MatchCt desc
, mt.Specification2
, mt.Family
, mc.Item
--
--
; --
'''SQL'''
#End If
End Function

Public Function sqlOf_GnsTubeHose( _
    Optional Diam As Double = 0 _
) As String
    sqlOf_GnsTubeHose = sqlOf_GnsTubeHose_v0_1(Diam)
End Function

Public Function sqlOf_GnsTubeHose_v0_1( _
    Optional Diam As Double = 0 _
) As String ', Matl As String, Dims As Variant
    ''' DON'T try to do anything with this yet!
    ''' see notes on where things are
    Dim txDiam As String
    
    If Diam > 0 Then
        txDiam = Join(Array( _
            "between", CStr(Diam - 0.01), _
            "and", CStr(Diam + 0.01) _
        ), " ")
    Else
        txDiam = "> 0.0"
    End If
    
    sqlOf_GnsTubeHose_v0_1 = _
        Replace(sqlTextLocal( _
        "sqlOf_GnsTubeHose_v0_1" _
        ), "%%DI%%", txDiam _
    ) '''
#If False Then
'''SQL'''
select i.Item,
i.Item + ' -- ' + i.Description1 Description
from vgMfiItems i
where i.Diameter %%DI%%
  and i.Specification1 in ('TUBE', 'HOSE')
  and i.Specification2 in ('HYDRAULIC', 'GREASELINE')
  and ISNULL(i.Specification16, '') = ''
order by i.Diameter
, i.Specification1 desc
, i.Specification2 desc
, i.Length desc
, i.Specification6 desc
, i.Item
--
--
; --
'''SQL'''
#End If
End Function

Public Function sqlOf_ASDF(Item As String) As String
    sqlOf_ASDF = Replace( _
        sqlTextLocal( _
        "sqlOf_ASDF" _
        ), "%%%", Item _
    )
#If False Then
'''SQL'''
-- SQL STATEMENT
with t as (
    select iType, bomStr
    from (values
        ('M', 51970), -- kNormalBOMStructure
        ('R', 51973)  -- kPurchasedBOMStructure
    ) ls(iType, bomStr)
)
select i.Item [Part Number]
, i.Family [Cost Center]
, i.Type
, t.bomStr
, b.ItemOrder Ord
, b.Item RM -- was Material
, m.Family MtFamily
, b.QuantityInConversionUnit RMQTY -- was Qty
, b.ConversionUnit RMUNIT -- was Unit
from vgIcoBillOfMaterials b
     inner join vgMfiItems i
       on b.Product = i.Item
     inner join vgMfiItems m
       on b.Item = m.Item
     inner join t
       on i.Type = t.iType
where b.Product = '%%%' -- 01-149
; --
'''SQL'''
#End If
End Function

Public Function sqlOf_03R4LC09_NOCOND() As String
#If False Then
'''SQL'''
-- SQL STATEMENT
-- 03R4LC09-NOCOND
Select
  H.Item,
  H.Description1 As Description,
  H.OptionPrice,
  I.Specification6 As Screen
from
  vgMfiItems As H Inner Join
  vgIcoBillOfMaterials B On H.Item = B.Product Inner Join
  vgMfiItems I On B.Item = I.Item
where
  IsNull(H.Specification16, '') = '' And
  H.Specification1 = 'BLENDER' And
  H.Specification2 = 'HOPPER' And
  H.Specification3 = 'RAISED4LC' And
  H.Specification5 In ('HOPPER', 'DRUM', 'COND') And
  H.Specification4 = '9' And
  I.Specification5 = 'SCREEN'
Order by
  H.Item
--
; --
'''SQL'''
#End If
    sqlOf_03R4LC09_NOCOND = sqlTextLocal( _
        "sqlOf_03R4LC09_NOCOND" _
    )
End Function

Public Function sqlOf_ERC_PTOSIZE() As String
'''SQL'''
'-- ERC-PTOSIZE
'select I.Item, I.Description1, I.OptionPrice, I.Specification7
', D.Item as PartsKit
'
'from vgMfiItems I
'inner join vgMfiItems D
'on  I.Specification1 = D.Specification1
'and I.Specification2 = D.Specification2
'and I.Specification4 = D.Specification4
'and I.Specification5 = D.Specification5
'-- and
'
'where I.Specification1 = 'SPREADER'
'and I.Specification2 ='PTO'
'and ISNULL(I.Specification3,'') = ''
'and ISNULL(D.Specification3,'') <> ''
'and I.Specification4 ='ALL'
'and I.Specification5 ='DRIVE'
'
'Order by Description1
'; --
'''SQL'''
    sqlOf_ERC_PTOSIZE = sqlTextLocal( _
        "sqlOf_ERC_PTOSIZE" _
    ) 'vbTextOfProcInDict
End Function

Public Function sqlOf_test2() As String
#If False Then
'''SQL'''
-- ERC-PTOSIZE
select I.Item, I.Description1, I.OptionPrice, I.Specification7
, D.Item as PartsKit

from vgMfiItems I
inner join vgMfiItems D
on  I.Specification1 = D.Specification1
and I.Specification2 = D.Specification2
and I.Specification4 = D.Specification4
and I.Specification5 = D.Specification5
-- and

where I.Specification1 = 'SPREADER'
and I.Specification2 ='PTO'
and ISNULL(I.Specification3,'') = ''
and ISNULL(D.Specification3,'') <> ''
and I.Specification4 ='ALL'
and I.Specification5 ='DRIVE'

Order by Description1
; --
'''SQL'''
#End If
    sqlOf_test2 = sqlTextLocal( _
        "sqlOf_test2" _
    )
'Debug.Print cnGnsDoyle().Execute(sqlOf_test2()).GetString
End Function

'''
''' END OF MODULE -- Add new VBA before this comment block.
'''
