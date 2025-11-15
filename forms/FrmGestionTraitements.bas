Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =6825
    ItemSuffix =13
    Left =5460
    Top =375
    Right =12630
    Bottom =2985
    TimerInterval =10
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x8e770cedd965e240
    End
    Caption ="DI 2003 - Traitements En Cours - DeltaInformatique 2003"
    DatasheetFontName ="Arial"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    AllowLayoutView =0
    PictureSizeMode =1
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            Height =1980
            Name ="Détail"
            Begin
                Begin Label
                    OverlapFlags =93
                    Top =285
                    Width =1695
                    Height =345
                    Name ="EtiFonLibelle"
                    Caption ="FONCTION :"
                    ControlTipText ="FONCTION :"
                End
                Begin Label
                    OverlapFlags =95
                    Top =630
                    Width =1680
                    Height =1065
                    Name ="EtiObjLibelle"
                    Caption ="OBJET :"
                    ControlTipText ="OBJET :"
                End
                Begin Label
                    OverlapFlags =95
                    Width =1695
                    Height =285
                    Name ="EtiTitLibelle"
                    Caption ="TITRE :"
                    ControlTipText ="TITRE :"
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =95
                    Top =1695
                    Width =6825
                    Height =285
                    Name ="BteProgressionFond"
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =87
                    Left =30
                    Top =1725
                    Width =0
                    Height =225
                    BackColor =16711680
                    Name ="BteProgression"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1695
                    Top =630
                    Width =5130
                    Height =1065
                    Name ="TxtObjValeur"
                    ControlTipText ="OBJET :"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1695
                    Top =300
                    Width =5115
                    Height =330
                    TabIndex =1
                    Name ="TxtFonValeur"
                    ControlTipText ="FONCTION :"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1695
                    Width =5115
                    Height =285
                    TabIndex =2
                    Name ="TxtTitValeur"
                    ControlTipText ="TITRE :"

                End
            End
        End
    End
End
CodeBehindForm
' See "FrmGestionTraitements.cls"
