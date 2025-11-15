Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =6240
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =2625
    Top =150
    Right =9210
    Bottom =5310
    RecSrcDt = Begin
        0xc811440b399fe140
    End
    Caption ="DI 2003 - Liste Des Fichiers - DeltaInformatique 2003"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnActivate ="[Event Procedure]"
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
        Begin Image
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
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            Height =4530
            Name ="Détail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =45
                    Top =3975
                    Width =1650
                    Height =480
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =-2147483640
                    Name ="CmdValider"
                    Caption ="Valider"
                    StatusBarText ="Valider"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Valider"

                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    Left =1137
                    Top =1185
                    Width =4701
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="TxtFichiersRepertoire"
                    StatusBarText ="Repertoire"
                    FontName ="Arial"
                    ControlTipText ="Repertoire"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =90
                            Top =1185
                            Width =1050
                            Height =240
                            ForeColor =-2147483640
                            Name ="EtiFichiersRepertoire"
                            Caption ="Emplacement"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    Left =1145
                    Top =1470
                    Width =4941
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="TxtExtension"
                    ControlSource ="=CmbFichiersListe.column(2)"
                    StatusBarText ="Extension"
                    FontName ="Arial"
                    ControlTipText ="Extension"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =90
                            Top =1470
                            Width =1050
                            Height =240
                            ForeColor =-2147483640
                            Name ="EtiFichiersExtension"
                            Caption ="Extension"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =93
                    Left =4530
                    Top =3975
                    Width =1605
                    Height =480
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =-2147483640
                    Name ="CmdFermer"
                    Caption ="Fermer"
                    StatusBarText ="Fermer"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Fermer"

                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =56
                    Top =563
                    Width =6097
                    Height =3902
                    Name ="RecFichiersPrincipal"
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =75
                    Top =75
                    Width =4710
                    Height =390
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483640
                    Name ="EtiFichiersIntitule"
                    Caption ="Liste Des Fichiers"
                    FontName ="Arial"
                    ControlTipText ="Liste Des Fichiers"
                End
                Begin Rectangle
                    SpecialEffect =1
                    OverlapFlags =255
                    Left =112
                    Top =1764
                    Width =5981
                    Height =2150
                    Name ="RecFichiersMessage"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    ColumnCount =2
                    ListWidth =4536
                    Left =1145
                    Top =915
                    Width =4941
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CmbFichiersPreference"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FicCode1, FicValeur FROM SelFichiersDetailler WHERE SelFichiersDetailler."
                        "FicCode2=[CmbFichiersListe] UNION SELECT TBLMACHINES.MacNom AS FicCode1, '' AS F"
                        "icValeur FROM TBLMACHINES GROUP BY TBLMACHINES.MacNom, '' HAVING TBLMACHINES.Mac"
                        "Nom not In (SELECT FicCode1 FROM SelFichiersDetailler WHERE SelFichiersDetailler"
                        ".FicCode2=[CmbFichiersListe];) UNION SELECT TBLMACHINES.MacUtilisateur AS FicCod"
                        "e1,'' As FicValeur FROM TBLMACHINES GROUP BY TBLMACHINES.MacUtilisateur,'' HAVIN"
                        "G TBLMACHINES.MacUtilisateur not In (SELECT FicCode1 FROM SelFichiersDetailler W"
                        "HERE SelFichiersDetailler.FicCode2=[CmbFichiersListe];) UNION SELECT TBLMACHINES"
                        ".MacDomaine AS FicCode1,'' As FicValeur FROM TBLMACHINES GROUP BY TBLMACHINES.Ma"
                        "cDomaine,'' HAVING TBLMACHINES.MacDomaine not In (SELECT FicCode1 FROM SelFichie"
                        "rsDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe];);"
                    ColumnWidths ="2268;2268"
                    StatusBarText ="Liste Des Preference"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyPress ="[Event Procedure]"
                    ControlTipText ="Liste Des Preferences"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =90
                            Top =915
                            Width =1050
                            Height =240
                            ForeColor =-2147483640
                            Name ="EtiFichiersPreference"
                            Caption ="Préférence"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    ColumnCount =4
                    ListWidth =2268
                    Left =1145
                    Top =630
                    Width =4941
                    BoundColumn =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CmbFichiersListe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SelFichiersDetailler.FicType, SelFichiersDetailler.FicType1, SelFichiersD"
                        "etailler.FicType2, SelFichiersDetailler.FicCode2 FROM SelFichiersDetailler WHERE"
                        " (((SelFichiersDetailler.FicType1)=\"FIC\")) GROUP BY SelFichiersDetailler.FicTy"
                        "pe, SelFichiersDetailler.FicType1, SelFichiersDetailler.FicType2, SelFichiersDet"
                        "ailler.FicCode2; "
                    ColumnWidths ="0;0;0;2268"
                    StatusBarText ="Liste Des Fichiers"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyPress ="[Event Procedure]"
                    ControlTipText ="Liste Des Fichiers"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =90
                            Top =615
                            Width =1050
                            Height =240
                            ForeColor =-2147483640
                            Name ="EtiFichiersNom"
                            Caption ="Nom"
                            FontName ="Arial"
                            ControlTipText ="Nom"
                        End
                    End
                End
                Begin ListBox
                    SpecialEffect =4
                    OverlapFlags =255
                    BorderWidth =3
                    ColumnCount =5
                    Left =135
                    Top =1815
                    Width =5880
                    Height =960
                    FontSize =10
                    TabIndex =6
                    BoundColumn =-1
                    ForeColor =255
                    Name ="LstFichiersManquant"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SelFichiersDetailler.FicType, SelFichiersDetailler.FicType1, SelFichiersD"
                        "etailler.FicType2, SelFichiersDetailler.FicCode2, Min(SelFichiersDetailler.FicVa"
                        "lide) AS FicValide FROM SelFichiersDetailler WHERE (((SelFichiersDetailler.FicTy"
                        "pe1)=\"FIC\")) GROUP BY SelFichiersDetailler.FicType, SelFichiersDetailler.FicTy"
                        "pe1, SelFichiersDetailler.FicType2, SelFichiersDetailler.FicCode2 HAVING (((Min("
                        "SelFichiersDetailler.FicValide))=0)); "
                    ColumnWidths ="0;567;567;2268;0"
                    StatusBarText ="Liste Des Fichiers Manquant"
                    FontName ="Arial"
                    ControlTipText ="Liste Des Fichiers Manquant"

                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =5835
                    Top =1170
                    Width =255
                    Height =255
                    TabIndex =7
                    Name ="CmdRepertoireFichier"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =2265
                    Top =3975
                    Width =1710
                    Height =480
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    ForeColor =-2147483640
                    Name ="CmdSupprimer"
                    Caption ="Supprimer"
                    StatusBarText ="Fermer"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Fermer"

                    Overlaps =1
                End
                Begin ListBox
                    SpecialEffect =4
                    OverlapFlags =247
                    BorderWidth =3
                    ColumnCount =5
                    Left =135
                    Top =2790
                    Width =5880
                    Height =1050
                    FontSize =10
                    TabIndex =9
                    BoundColumn =-1
                    ForeColor =16711680
                    Name ="LstFichiersValide"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SelFichiersDetailler.FicType, SelFichiersDetailler.FicType1, SelFichiersD"
                        "etailler.FicType2, SelFichiersDetailler.FicCode2, Min(SelFichiersDetailler.FicVa"
                        "lide) AS FicValide FROM SelFichiersDetailler WHERE (((SelFichiersDetailler.FicTy"
                        "pe1)=\"FIC\")) GROUP BY SelFichiersDetailler.FicType, SelFichiersDetailler.FicTy"
                        "pe1, SelFichiersDetailler.FicType2, SelFichiersDetailler.FicCode2 HAVING (((Min("
                        "SelFichiersDetailler.FicValide))=-1)) ORDER BY SelFichiersDetailler.FicType, Sel"
                        "FichiersDetailler.FicCode2, Min(SelFichiersDetailler.FicValide) DESC; "
                    ColumnWidths ="0;567;567;2268;0"
                    StatusBarText ="Liste Des Fichiers Valide"
                    FontName ="Arial"
                    ControlTipText ="Liste Des Fichiers Valide"

                End
                Begin CustomControl
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5669
                    Top =56
                    Width =480
                    Height =480
                    AutoActivate =1
                    TabIndex =10
                    Name ="CdgFichiers"
                    OLEClass ="CommonDialog"
                    Class ="MSComDlg.CommonDialog.1"

                End
            End
        End
    End
End
CodeBehindForm
' See "FrmGestionFichiers.cls"
