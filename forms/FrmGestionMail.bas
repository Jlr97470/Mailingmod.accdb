Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =6365
    DatasheetFontHeight =10
    ItemSuffix =38
    Left =6270
    Top =1605
    Right =12990
    Bottom =7440
    RecSrcDt = Begin
        0xc811440b399fe140
    End
    Caption ="DI 2003 - Liste Des Mails - DeltaInformatique 2003"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
            Height =5175
            Name ="Détail"
            Begin
                Begin Rectangle
                    OverlapFlags =93
                    Left =56
                    Top =563
                    Width =6097
                    Height =3902
                    Name ="RecEMailPrincipal"
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
                    BackColor =0
                    Name ="EtiMailsIntitule"
                    Caption ="Liste Des Mails"
                    FontName ="Arial"
                    ControlTipText ="Liste Des Mails"
                End
                Begin Rectangle
                    SpecialEffect =1
                    OverlapFlags =223
                    Left =142
                    Top =2589
                    Width =5951
                    Height =1820
                    Name ="RecEMailFichiersAttacher"
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    SpecialEffect =4
                    OverlapFlags =215
                    BorderWidth =3
                    ColumnCount =2
                    Left =195
                    Top =2655
                    Width =5820
                    Height =1695
                    FontSize =10
                    TabIndex =3
                    BoundColumn =-1
                    ForeColor =16711680
                    Name ="LstEMailFichiersAttacher"
                    RowSourceType ="Value List"
                    ColumnWidths ="2268;3402"
                    StatusBarText ="Liste Des Fichiers Attacher"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Liste Des Fichiers Attacher"

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
                    TabIndex =4
                    Name ="CdgFichiers"
                    OLEClass ="CommonDialog"
                    Class ="MSComDlg.CommonDialog.1"

                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2000
                    Top =630
                    Width =4086
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="TxtEMailExpediteur"
                    StatusBarText ="Expediteur"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyPress ="[Event Procedure]"
                    ControlTipText ="Expediteur"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =90
                            Top =615
                            Width =1830
                            Height =240
                            BackColor =0
                            Name ="EtiEMailExpediteur"
                            Caption ="Expediteur"
                            FontName ="Arial"
                            ControlTipText ="Expediteur"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2000
                    Top =915
                    Width =4086
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="TxtEMailCopieCacher"
                    StatusBarText ="Copie Cacher"
                    FontName ="Arial"
                    OnKeyPress ="[Event Procedure]"
                    ControlTipText ="Copie Cacher"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =90
                            Top =915
                            Width =1830
                            Height =240
                            BackColor =0
                            Name ="EtiEMailCopieCacher"
                            Caption ="Copie Cachée"
                            FontName ="Arial"
                            ControlTipText ="Copie Cacher"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =5580
                    Top =4545
                    Width =591
                    Height =591
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="CmdFermer"
                    Caption ="Fermer Le Formulaire"
                    StatusBarText ="Fermer Le Formulaire"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadadadadadaadad00adad00adaddadad00ad00adada ,
                        0xadadad0000adadaddadadad00adadadaadadad0000adadaddadad00ad00adada ,
                        0xadad00adad00adaddadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="System"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Fermer Le Formulaire"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =75
                    Top =4560
                    Width =1101
                    Height =591
                    FontWeight =700
                    TabIndex =6
                    ForeColor =-2147483630
                    Name ="CmdExporteMailing"
                    Caption ="ENVOYER"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Exécuter Excel"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4536
                    Left =2000
                    Top =1200
                    Width =4086
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CmbEMailChampDestination"
                    RowSourceType ="Value List"
                    RowSource ="CiviliteIndiv;;NomIndiv;;PrenomIndiv;;FonctionIndiv;;EstDirigeant;;EstIC;;LoginI"
                        "ndiv;;PassIndiv;;EmailIndiv;"
                    ColumnWidths ="2268;2268"
                    StatusBarText ="Champ Destination"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Champ Destination"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =90
                            Top =1200
                            Width =1845
                            Height =240
                            BackColor =0
                            Name ="EtiEMailChampDestination"
                            Caption ="Champ Destination"
                            FontName ="Arial"
                            ControlTipText ="Champ Destination"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4536
                    Left =2000
                    Top =1485
                    Width =4086
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CmbEMailFichierBody"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SelFichiersDetailler.FicCode2, SelFichiersDetailler.FicValeur FROM SelFic"
                        "hiersDetailler WHERE (((SelFichiersDetailler.FicValide)=True) AND ((SelFichiersD"
                        "etailler.FicType2)='HTM')); "
                    ColumnWidths ="2268;2268"
                    StatusBarText ="Fichier Body"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Fichier Body"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =90
                            Top =1485
                            Width =1845
                            Height =240
                            BackColor =0
                            Name ="EtiEMailFichierBody"
                            Caption ="Fichier Body"
                            FontName ="Arial"
                            ControlTipText ="Fichier Body"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3585
                    Top =2085
                    Width =2475
                    Height =405
                    FontWeight =700
                    TabIndex =8
                    Name ="CmdAjouterFichierAttacher"
                    Caption ="Ajouter un fichier attacher"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Dernier enregistrement"

                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2000
                    Top =1770
                    Width =4086
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="TxtEMailSujet"
                    StatusBarText ="Copie Cacher"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Copie Cacher"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =90
                            Top =1770
                            Width =1830
                            Height =240
                            BackColor =0
                            Name ="EtiEMailSujet"
                            Caption ="Sujet"
                            FontName ="Arial"
                            ControlTipText ="Copie Cacher"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =165
                    Top =2190
                    Width =185
                    Height =195
                    TabIndex =10
                    Name ="ChkEmailSimulation"
                    DefaultValue ="True"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =480
                            Top =2160
                            Width =810
                            Height =240
                            Name ="EtiEmailSimulation"
                            Caption ="Simulation"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmGestionMail.cls"
