Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    PictureTiling = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =5
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8505
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =4995
    Top =1125
    Right =13875
    Bottom =5970
    PaintPalette = Begin
        0x000301000000000000000000
    End
    RecSrcDt = Begin
        0x1310e2049a97e240
    End
    RecordSource ="SELECT TBLPARAMETRES.* FROM TBLPARAMETRES; "
    Caption ="DI 2003 - Listes Des Parametres - DeltaInformatique 2003"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnActivate ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            SpecialEffect =1
            FontWeight =700
            BackColor =12632256
            ForeColor =128
            FontName ="Arial"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
            Width =1701
        End
        Begin Image
            SpecialEffect =3
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            ForeColor =128
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderWidth =3
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BackStyle =1
            BorderLineStyle =0
            Width =1701
            Height =1701
            BackColor =12632256
        End
        Begin BoundObjectFrame
            SpecialEffect =3
            BorderLineStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
            BorderColor =12632256
            BackColor =12632256
        End
        Begin TextBox
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =12632256
            BorderColor =12632256
            FontName ="Arial"
        End
        Begin ListBox
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =12632256
            BorderColor =12632256
            FontName ="Arial"
        End
        Begin ComboBox
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =12632256
            BorderColor =12632256
            FontName ="Arial"
        End
        Begin Subform
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin UnboundObjectFrame
            SpecialEffect =3
            BackStyle =0
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            ForeColor =128
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            Width =5103
            Height =3402
            FontWeight =700
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =885
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    SpecialEffect =4
                    BackStyle =0
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =93
                    TextAlign =1
                    Top =570
                    Width =1125
                    Height =300
                    FontSize =10
                    ForeColor =16711680
                    Name ="EtiParType"
                    Caption ="Type"
                    Tag ="Type"
                    ControlTipText ="Type"
                End
                Begin Label
                    SpecialEffect =4
                    BackStyle =0
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =95
                    TextAlign =1
                    Left =1133
                    Top =566
                    Width =3120
                    Height =300
                    FontSize =10
                    ForeColor =16711680
                    Name ="EtiParCode"
                    Caption ="Code"
                    Tag ="Code"
                    ControlTipText ="Code"
                End
                Begin Label
                    SpecialEffect =4
                    BackStyle =0
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =95
                    TextAlign =1
                    Left =4260
                    Top =570
                    Width =3540
                    Height =300
                    FontSize =10
                    ForeColor =16711680
                    Name ="EtiParValeur"
                    Caption ="Valeur"
                    Tag ="Valeur"
                    ControlTipText ="Valeur"
                End
                Begin Label
                    SpecialEffect =4
                    BackStyle =0
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =2
                    Left =2835
                    Top =135
                    Width =2850
                    Height =300
                    FontSize =10
                    ForeColor =16711680
                    Name ="EtiParametres"
                    Caption ="Listes Des Paramétres"
                    Tag ="Listes Des Parametres"
                    ControlTipText ="Listes Des Parametres"
                End
                Begin Label
                    SpecialEffect =4
                    BackStyle =0
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =87
                    TextAlign =1
                    Left =7830
                    Top =570
                    Width =675
                    Height =300
                    FontSize =10
                    ForeColor =16711680
                    Name ="EtiParValide"
                    Caption ="Valide"
                    Tag ="Valeur"
                    ControlTipText ="Valeur"
                End
            End
        End
        Begin Section
            Height =255
            Name ="Détail"
            Begin
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =93
                    ListWidth =2268
                    Left =15
                    Top =15
                    Width =1131
                    Name ="LstParType"
                    ControlSource ="ParType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT TBLPARAMETRES.ParType FROM TBLPARAMETRES; "
                    ColumnWidths ="2268"
                    StatusBarText ="Parametre Type"
                    OnKeyPress ="[Event Procedure]"
                    ControlTipText ="Parametre Type"

                End
                Begin TextBox
                    OverlapFlags =95
                    Left =1140
                    Width =3111
                    TabIndex =1
                    Name ="TxtParCode"
                    ControlSource ="ParCode"
                    StatusBarText ="Parametre Code"
                    OnKeyPress ="[Event Procedure]"
                    ControlTipText ="Parametre Code"

                End
                Begin TextBox
                    OverlapFlags =87
                    Left =4260
                    Width =3516
                    TabIndex =2
                    Name ="TxtParValeur"
                    ControlSource ="ParValeur"
                    StatusBarText ="Parametre Valeur"
                    OnKeyPress ="[Event Procedure]"
                    ControlTipText ="Parametre Valeur"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8070
                    Top =30
                    Width =200
                    Height =195
                    ColumnWidth =2475
                    TabIndex =3
                    Name ="CbxParValide"
                    ControlSource ="ParValide"
                    StatusBarText ="Parametres Valide (Oui/Non)"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="PiedFormulaire"
        End
    End
End
CodeBehindForm
' See "FrmGestionParametres.cls"
