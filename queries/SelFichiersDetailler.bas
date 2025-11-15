Operation =1
Option =0
Begin InputTables
    Name ="TBLFICHIERS"
End
Begin OutputColumns
    Expression ="TBLFICHIERS.FicType"
    Alias ="FicType1"
    Expression ="Left([FicType],3)"
    Alias ="FicType2"
    Expression ="Mid([FicType],4)"
    Expression ="TBLFICHIERS.FicCode"
    Alias ="FicCode1"
    Expression ="Left([FicCode],InStr([FicCode],\"=\")-1)"
    Alias ="FicCode2"
    Expression ="Right([FicCode],Len([FicCode])-InStr([FicCode],\"=\"))"
    Expression ="TBLFICHIERS.FicValeur"
    Expression ="TBLFICHIERS.FicValide"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbText "Description" ="Selection Des Fichiers Detailler"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FicType1"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FicType2"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FicCode1"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FicCode2"
        dbInteger "ColumnWidth" ="4185"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBLFICHIERS.FicValide"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBLFICHIERS.FicValeur"
        dbInteger "ColumnWidth" ="7905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBLFICHIERS.FicType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =45
    Top =94
    Right =1081
    Bottom =687
    Left =-1
    Top =-1
    Right =1012
    Bottom =415
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="TBLFICHIERS"
        Name =""
    End
End
