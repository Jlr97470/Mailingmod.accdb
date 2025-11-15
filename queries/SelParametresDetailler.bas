Operation =1
Option =0
Begin InputTables
    Name ="TBLPARAMETRES"
End
Begin OutputColumns
    Expression ="TBLPARAMETRES.ParType"
    Alias ="ParType1"
    Expression ="Left([ParType],3)"
    Alias ="ParType2"
    Expression ="Right([ParType],3)"
    Expression ="TBLPARAMETRES.ParCode"
    Alias ="ParCode1"
    Expression ="Left([ParCode],InStr([ParCode],\"=\")-1)"
    Alias ="ParCode2"
    Expression ="Left(RemplaceChr([ParCode],[ParCode1] & \"=\",\"\"),InStr(RemplaceChr([ParCode],"
        "[ParCode1] & \"=\",\"\"),\"=\")-1)"
    Alias ="ParCode3"
    Expression ="Left(RemplaceChr([ParCode],[ParCode1] & \"=\" & [ParCode2] & \"=\",\"\"),InStr(R"
        "emplaceChr([ParCode],[ParCode1] & \"=\" & [ParCode2] & \"=\",\"\"),\"=\")-1)"
    Alias ="ParCode4"
    Expression ="Left(RemplaceChr([ParCode],[ParCode1] & \"=\" & [ParCode2] & \"=\" & [ParCode3] "
        "& \"=\",\"\"),InStr(RemplaceChr([ParCode],[ParCode1] & \"=\" & [ParCode2] & \"=\""
        " & [ParCode3] & \"=\",\"\"),\"=\")-1)"
    Alias ="ParCode5"
    Expression ="RemplaceChr([ParCode],[ParCode1] & \"=\" & [ParCode2] & \"=\" & [ParCode3] & \"="
        "\" & [ParCode4] & \"=\",\"\")"
    Expression ="TBLPARAMETRES.ParValeur"
    Expression ="TBLPARAMETRES.ParValide"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbText "Description" ="Selection Des Parametres Detailler"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="TBLPARAMETRES.ParType"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="TBLPARAMETRES.ParCode"
        dbInteger "ColumnWidth" ="9750"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="TBLPARAMETRES.ParValeur"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ParType1"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="ParType2"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="ParCode1"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="ParCode2"
        dbInteger "ColumnWidth" ="3570"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="ParCode3"
        dbInteger "ColumnWidth" ="945"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ParCode4"
        dbInteger "ColumnWidth" ="945"
        dbInteger "ColumnOrder" ="0"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ParCode5"
        dbInteger "ColumnWidth" ="5610"
        dbInteger "ColumnOrder" ="0"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =40
    Top =22
    Right =1258
    Bottom =327
    Left =-1
    Top =-1
    Right =1207
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="TBLPARAMETRES"
        Name =""
    End
End
