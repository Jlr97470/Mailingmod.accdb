Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =220
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8475
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =4635
    Top =885
    Right =13470
    Bottom =6645
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x30d5b8a71595e240
    End
    Caption ="DI 2003 - Notes - DeltaInformatique 2003"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =5130
            BackColor =-2147483633
            Name ="Détail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =8400
                    Height =5070
                    Name ="TxtNotes"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =30
                            Width =645
                            Height =240
                            Name ="Étiquette1"
                            Caption ="Texte0:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmNotes.cls"
