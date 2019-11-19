VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Begin VB.Form frmINV_OS 
   Caption         =   "Фильтр для Карточка основного средства"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Tag             =   "Card.ICO"
   Begin MSComctlLib.TabStrip ts 
      Height          =   1500
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   330
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Отказ от задания фильтра"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   750
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Применить фильтр"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   750
   End
   Begin MTZ_PANEL.ScrolledWindow PanelfGroup 
      Height          =   1000
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   1773
      Begin MSComCtl2.DTPicker dtpINVOS_SROK_RecalcDate_LE 
         Height          =   300
         Left            =   12900
         TabIndex        =   81
         ToolTipText     =   "Дата следующего пересчета срока по"
         Top             =   405
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74514435
         CurrentDate     =   40142
      End
      Begin VB.CheckBox lblINVOS_SROK_RecalcDate_LE 
         Caption         =   "Дата следующего пересчета срока по:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12900
         TabIndex        =   80
         Top             =   75
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpINVOS_SROK_RecalcDate_GE 
         Height          =   300
         Left            =   9750
         TabIndex        =   79
         ToolTipText     =   "Дата следующего пересчета срока C"
         Top             =   6240
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74514435
         CurrentDate     =   40142
      End
      Begin VB.CheckBox lblINVOS_SROK_RecalcDate_GE 
         Caption         =   "Дата следующего пересчета срока C:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9750
         TabIndex        =   78
         Top             =   5910
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_CODE_VisibleCode 
         Height          =   300
         Left            =   9750
         MaxLength       =   255
         TabIndex        =   77
         ToolTipText     =   "Читаемый код"
         Top             =   5535
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_CODE_VisibleCode 
         Caption         =   "Читаемый код:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9750
         TabIndex        =   76
         Top             =   5205
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_PLACE_Info 
         Height          =   1200
         Left            =   9750
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   75
         ToolTipText     =   "Описание"
         Top             =   3930
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_PLACE_Info 
         Caption         =   "Описание:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9750
         TabIndex        =   74
         Top             =   3600
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_TheOwner 
         Height          =   300
         Left            =   12300
         TabIndex        =   73
         Tag             =   "refopen.ico"
         ToolTipText     =   "Владелец"
         Top             =   3225
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtINVOS_PLACE_TheOwner 
         Height          =   300
         Left            =   9750
         Locked          =   -1  'True
         TabIndex        =   72
         ToolTipText     =   "Владелец"
         Top             =   3225
         Width           =   2550
      End
      Begin VB.CheckBox lblINVOS_PLACE_TheOwner 
         Caption         =   "Владелец:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9750
         TabIndex        =   71
         Top             =   2895
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_PLACE_WorkPlaceNum 
         Height          =   300
         Left            =   9750
         MaxLength       =   10
         TabIndex        =   70
         ToolTipText     =   "Номер рабочего места"
         Top             =   2520
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_PLACE_WorkPlaceNum 
         Caption         =   "Номер рабочего места:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9750
         TabIndex        =   69
         Top             =   2190
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_PLACE_Room 
         Height          =   300
         Left            =   9750
         MaxLength       =   10
         TabIndex        =   68
         ToolTipText     =   "Кабинет"
         Top             =   1815
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_PLACE_Room 
         Caption         =   "Кабинет:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9750
         TabIndex        =   67
         Top             =   1485
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_PLACE_Flow 
         Height          =   300
         Left            =   9750
         MaxLength       =   10
         TabIndex        =   66
         ToolTipText     =   "Этаж"
         Top             =   1110
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_PLACE_Flow 
         Caption         =   "Этаж:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9750
         TabIndex        =   65
         Top             =   780
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_Otdel 
         Height          =   300
         Left            =   12300
         TabIndex        =   64
         Tag             =   "refopen.ico"
         ToolTipText     =   "Отдел"
         Top             =   405
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtINVOS_PLACE_Otdel 
         Height          =   300
         Left            =   9750
         Locked          =   -1  'True
         TabIndex        =   63
         ToolTipText     =   "Отдел"
         Top             =   405
         Width           =   2550
      End
      Begin VB.CheckBox lblINVOS_PLACE_Otdel 
         Caption         =   "Отдел:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9750
         TabIndex        =   62
         Top             =   75
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_Uprav 
         Height          =   300
         Left            =   9150
         TabIndex        =   61
         Tag             =   "refopen.ico"
         ToolTipText     =   "Управление"
         Top             =   6240
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtINVOS_PLACE_Uprav 
         Height          =   300
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   60
         ToolTipText     =   "Управление"
         Top             =   6240
         Width           =   2550
      End
      Begin VB.CheckBox lblINVOS_PLACE_Uprav 
         Caption         =   "Управление:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   59
         Top             =   5910
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_DIrection 
         Height          =   300
         Left            =   9150
         TabIndex        =   58
         Tag             =   "refopen.ico"
         ToolTipText     =   "Дирекция"
         Top             =   5535
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtINVOS_PLACE_DIrection 
         Height          =   300
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   57
         ToolTipText     =   "Дирекция"
         Top             =   5535
         Width           =   2550
      End
      Begin VB.CheckBox lblINVOS_PLACE_DIrection 
         Caption         =   "Дирекция:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   56
         Top             =   5205
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_PLACE_ComplNumber 
         Height          =   300
         Left            =   6600
         MaxLength       =   30
         TabIndex        =   55
         ToolTipText     =   "Номер комплекта"
         Top             =   4830
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_PLACE_ComplNumber 
         Caption         =   "Номер комплекта:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   54
         Top             =   4500
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_TheHouse 
         Height          =   300
         Left            =   9150
         TabIndex        =   53
         Tag             =   "refopen.ico"
         ToolTipText     =   "Здание"
         Top             =   4125
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtINVOS_PLACE_TheHouse 
         Height          =   300
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   52
         ToolTipText     =   "Здание"
         Top             =   4125
         Width           =   2550
      End
      Begin VB.CheckBox lblINVOS_PLACE_TheHouse 
         Caption         =   "Здание:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   51
         Top             =   3795
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_MatOtv 
         Height          =   300
         Left            =   9150
         TabIndex        =   50
         Tag             =   "refopen.ico"
         ToolTipText     =   "Матерально отв."
         Top             =   3420
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtINVOS_PLACE_MatOtv 
         Height          =   300
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   49
         ToolTipText     =   "Матерально отв."
         Top             =   3420
         Width           =   2550
      End
      Begin VB.CheckBox lblINVOS_PLACE_MatOtv 
         Caption         =   "Матерально отв.:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   48
         Top             =   3090
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_TechFilePath 
         Height          =   300
         Left            =   6600
         MaxLength       =   255
         TabIndex        =   47
         ToolTipText     =   "Путь к файлу с ТИ"
         Top             =   2715
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_INFO_TechFilePath 
         Caption         =   "Путь к файлу с ТИ:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   46
         Top             =   2385
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_Info 
         Height          =   1200
         Left            =   6600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         ToolTipText     =   "Описание"
         Top             =   1110
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_INFO_Info 
         Caption         =   "Описание:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   44
         Top             =   780
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpINVOS_INFO_ActivateDate_LE 
         Height          =   300
         Left            =   6600
         TabIndex        =   43
         ToolTipText     =   "Дата ввода в эксп. по"
         Top             =   405
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74514435
         CurrentDate     =   40142
      End
      Begin VB.CheckBox lblINVOS_INFO_ActivateDate_LE 
         Caption         =   "Дата ввода в эксп. по:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   42
         Top             =   75
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpINVOS_INFO_ActivateDate_GE 
         Height          =   300
         Left            =   3450
         TabIndex        =   41
         ToolTipText     =   "Дата ввода в эксп. C"
         Top             =   6045
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74514435
         CurrentDate     =   40142
      End
      Begin VB.CheckBox lblINVOS_INFO_ActivateDate_GE 
         Caption         =   "Дата ввода в эксп. C:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   40
         Top             =   5715
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_SrokOI_LE 
         Height          =   300
         Left            =   3450
         MaxLength       =   15
         TabIndex        =   39
         ToolTipText     =   "Остаточный срок ПИ <="
         Top             =   5340
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_SrokOI_LE 
         Caption         =   "Остаточный срок ПИ <=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   38
         Top             =   5010
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_SrokOI_GE 
         Height          =   300
         Left            =   3450
         MaxLength       =   15
         TabIndex        =   37
         ToolTipText     =   "Остаточный срок ПИ >="
         Top             =   4635
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_SrokOI_GE 
         Caption         =   "Остаточный срок ПИ >=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   36
         Top             =   4305
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_SrokFI_LE 
         Height          =   300
         Left            =   3450
         MaxLength       =   15
         TabIndex        =   35
         ToolTipText     =   "Срок ФИ <="
         Top             =   3930
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_SrokFI_LE 
         Caption         =   "Срок ФИ <=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   34
         Top             =   3600
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_SrokFI_GE 
         Height          =   300
         Left            =   3450
         MaxLength       =   15
         TabIndex        =   33
         ToolTipText     =   "Срок ФИ >="
         Top             =   3225
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_SrokFI_GE 
         Caption         =   "Срок ФИ >=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   32
         Top             =   2895
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_SrokPI_LE 
         Height          =   300
         Left            =   3450
         MaxLength       =   15
         TabIndex        =   31
         ToolTipText     =   "Срок ПИ <="
         Top             =   2520
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_SrokPI_LE 
         Caption         =   "Срок ПИ <=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   30
         Top             =   2190
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_SrokPI_GE 
         Height          =   300
         Left            =   3450
         MaxLength       =   15
         TabIndex        =   29
         ToolTipText     =   "Срок ПИ >="
         Top             =   1815
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_SrokPI_GE 
         Caption         =   "Срок ПИ >=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   28
         Top             =   1485
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_TheCost_LE 
         Height          =   300
         Left            =   3450
         MaxLength       =   27
         TabIndex        =   27
         ToolTipText     =   "Cтоимость <="
         Top             =   1110
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_TheCost_LE 
         Caption         =   "Cтоимость <=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   26
         Top             =   780
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_TheCost_GE 
         Height          =   300
         Left            =   3450
         MaxLength       =   27
         TabIndex        =   25
         ToolTipText     =   "Cтоимость >="
         Top             =   405
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_TheCost_GE 
         Caption         =   "Cтоимость >=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   24
         Top             =   75
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_InLineNum_LE 
         Height          =   300
         Left            =   300
         MaxLength       =   15
         TabIndex        =   23
         ToolTipText     =   "Номер в партии <="
         Top             =   6045
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_InLineNum_LE 
         Caption         =   "Номер в партии <=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   22
         Top             =   5715
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_InLineNum_GE 
         Height          =   300
         Left            =   300
         MaxLength       =   15
         TabIndex        =   21
         ToolTipText     =   "Номер в партии >="
         Top             =   5340
         Width           =   1800
      End
      Begin VB.CheckBox lblINVOS_INFO_InLineNum_GE 
         Caption         =   "Номер в партии >=:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   20
         Top             =   5010
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_CardNum 
         Height          =   300
         Left            =   300
         MaxLength       =   30
         TabIndex        =   19
         ToolTipText     =   "Номер карточки учета"
         Top             =   4635
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_INFO_CardNum 
         Caption         =   "Номер карточки учета:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   18
         Top             =   4305
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_INVNum 
         Height          =   300
         Left            =   300
         MaxLength       =   30
         TabIndex        =   17
         ToolTipText     =   "Инвентарный номер"
         Top             =   3930
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_INFO_INVNum 
         Caption         =   "Инвентарный номер:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   16
         Top             =   3600
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_ShortName 
         Height          =   300
         Left            =   300
         MaxLength       =   100
         TabIndex        =   15
         ToolTipText     =   "Краткое наименование"
         Top             =   3225
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_INFO_ShortName 
         Caption         =   "Краткое наименование:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   14
         Top             =   2895
         Width           =   3000
      End
      Begin VB.TextBox txtINVOS_INFO_Name 
         Height          =   300
         Left            =   300
         MaxLength       =   255
         TabIndex        =   13
         ToolTipText     =   "Наименование"
         Top             =   2520
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_INFO_Name 
         Caption         =   "Наименование:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   12
         Top             =   2190
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdINVOS_INFO_OSType 
         Height          =   300
         Left            =   2850
         TabIndex        =   11
         Tag             =   "refopen.ico"
         ToolTipText     =   "Группа ОС"
         Top             =   1815
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtINVOS_INFO_OSType 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Группа ОС"
         Top             =   1815
         Width           =   2550
      End
      Begin VB.CheckBox lblINVOS_INFO_OSType 
         Caption         =   "Группа ОС:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   9
         Top             =   1485
         Width           =   3000
      End
      Begin VB.ComboBox cmbINVOS_INFO_IsMaterial 
         Height          =   315
         Left            =   300
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Материал"
         Top             =   1110
         Width           =   3000
      End
      Begin VB.CheckBox lblINVOS_INFO_IsMaterial 
         Caption         =   "Материал:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   7
         Top             =   780
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdINVOS_INFO_TheOrg 
         Height          =   300
         Left            =   2850
         TabIndex        =   6
         Tag             =   "refopen.ico"
         ToolTipText     =   "На учете в "
         Top             =   405
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtINVOS_INFO_TheOrg 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "На учете в "
         Top             =   405
         Width           =   2550
      End
      Begin VB.CheckBox lblINVOS_INFO_TheOrg 
         Caption         =   "На учете в :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   4
         Top             =   75
         Width           =   3000
      End
   End
   Begin VB.Menu mnuCtl 
      Caption         =   "mnuCtl"
      Visible         =   0   'False
      Begin VB.Menu mnuSetup 
         Caption         =   "Настройка"
      End
   End
End
Attribute VB_Name = "frmINV_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Item As Object
Public OK As Boolean
Private OnInit As Boolean
Public Event Changed()
Private TSCustom As MTZ_CUSTOMTAB.TabStripCustomizer







Private Sub cmdOK_Click()
    On Error Resume Next
    OK = True
    Me.Hide
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    OK = False
    Me.Hide
End Sub
Public Sub Init(ObjItem As Object)
 Set Item = ObjItem
 If Item Is Nothing Then Set Item = MyUser.Application
 TInit
End Sub
Private Sub Form_Load()
  On Error Resume Next
  Dim ff As Long, buf As String
  LoadFromSkin Me
  ts.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdOK.Height
  cmdOK.Move Me.ScaleWidth - 110 * Screen.TwipsPerPixelX, Me.ScaleHeight - cmdOK.Height, cmdOK.Width, cmdOK.Height
  cmdCancel.Move Me.ScaleWidth - 55 * Screen.TwipsPerPixelX, Me.ScaleHeight - cmdOK.Height, cmdCancel.Width, cmdOK.Height
  If Item Is Nothing Then Init MyUser.Application
End Sub
Private Sub Form_Unload(cancel As Integer)
  On Error Resume Next
  Set Item = Nothing
  Set TSCustom = Nothing
  SaveToSkin Me
  Exit Sub
bye:
  MsgBox Err.Description, vbOKOnly
  cancel = -1
End Sub
Private Sub Form_Resize()
 If Me.WindowState = 1 Then Exit Sub
 On Error Resume Next
  ts.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdOK.Height
  cmdOK.Move Me.ScaleWidth - 110 * Screen.TwipsPerPixelX, Me.ScaleHeight - cmdOK.Height, cmdOK.Width, cmdOK.Height
  cmdCancel.Move Me.ScaleWidth - 55 * Screen.TwipsPerPixelX, Me.ScaleHeight - cmdOK.Height, cmdCancel.Width, cmdOK.Height
  ts_click
End Sub
Private Sub mnuSetup_Click()
TSCustom.Setup ts
End Sub
Private Sub ts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 And Shift = 0 Then
    PopupMenu mnuCtl
  End If
End Sub
Private Sub ts_click()
  On Error Resume Next
  PanelfGroup.Visible = False

   Select Case ts.SelectedItem.Key
   Case "fGroup"
     With PanelfGroup
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     End With
     End Select
End Sub
Private Sub TInit()
  On Error Resume Next
  Dim ff As Long, buf As String

ts.Tabs.Item(1).Caption = "Общие сведения"
ts.Tabs.Item(1).Key = "fGroup"
PanelfGroupInit
  Set TSCustom = New MTZ_CUSTOMTAB.TabStripCustomizer
  TSCustom.Init ts, "INV_OS", "fctlINV_OS"
  TSCustom.SetupFromRegistry ts
 ts_click
End Sub


Private Sub Changing()
If OnInit Then Exit Sub
 RaiseEvent Changed
End Sub
Private Sub txtINVOS_INFO_TheOrg_Change()
  Changing
End Sub
Private Sub cmdINVOS_INFO_TheOrg_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_ORG", id, brief) Then
          txtINVOS_INFO_TheOrg.Tag = Left(id, 38)
          txtINVOS_INFO_TheOrg = brief
        End If
End Sub
Private Sub cmbINVOS_INFO_IsMaterial_Click()
  On Error Resume Next
  Changing
End Sub
Private Sub txtINVOS_INFO_OSType_Change()
  Changing
End Sub
Private Sub cmdINVOS_INFO_OSType_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_OSTYPE", id, brief) Then
          txtINVOS_INFO_OSType.Tag = Left(id, 38)
          txtINVOS_INFO_OSType = brief
        End If
End Sub
Private Sub txtINVOS_INFO_Name_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_ShortName_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_INVNum_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_CardNum_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_InLineNum_GE_Validate(cancel As Boolean)
If txtINVOS_INFO_InLineNum_GE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_InLineNum_GE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_InLineNum_GE.Text) <> CLng(val(txtINVOS_INFO_InLineNum_GE.Text)) Then
     cancel = True
     MsgBox "Ожидалось целое число", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_InLineNum_GE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_InLineNum_GE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_InLineNum_LE_Validate(cancel As Boolean)
If txtINVOS_INFO_InLineNum_LE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_InLineNum_LE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_InLineNum_LE.Text) <> CLng(val(txtINVOS_INFO_InLineNum_LE.Text)) Then
     cancel = True
     MsgBox "Ожидалось целое число", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_InLineNum_LE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_InLineNum_LE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_TheCost_GE_Validate(cancel As Boolean)
If txtINVOS_INFO_TheCost_GE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_TheCost_GE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_TheCost_GE.Text) < -922337203685478# Or val(txtINVOS_INFO_TheCost_GE.Text) > 922337203685478# Then
     cancel = True
     MsgBox "Значение вне допустимого диапазона", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_TheCost_GE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_TheCost_GE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_TheCost_LE_Validate(cancel As Boolean)
If txtINVOS_INFO_TheCost_LE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_TheCost_LE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_TheCost_LE.Text) < -922337203685478# Or val(txtINVOS_INFO_TheCost_LE.Text) > 922337203685478# Then
     cancel = True
     MsgBox "Значение вне допустимого диапазона", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_TheCost_LE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_TheCost_LE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_SrokPI_GE_Validate(cancel As Boolean)
If txtINVOS_INFO_SrokPI_GE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_SrokPI_GE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_SrokPI_GE.Text) <> CLng(val(txtINVOS_INFO_SrokPI_GE.Text)) Then
     cancel = True
     MsgBox "Ожидалось целое число", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_SrokPI_GE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_SrokPI_GE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_SrokPI_LE_Validate(cancel As Boolean)
If txtINVOS_INFO_SrokPI_LE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_SrokPI_LE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_SrokPI_LE.Text) <> CLng(val(txtINVOS_INFO_SrokPI_LE.Text)) Then
     cancel = True
     MsgBox "Ожидалось целое число", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_SrokPI_LE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_SrokPI_LE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_SrokFI_GE_Validate(cancel As Boolean)
If txtINVOS_INFO_SrokFI_GE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_SrokFI_GE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_SrokFI_GE.Text) <> CLng(val(txtINVOS_INFO_SrokFI_GE.Text)) Then
     cancel = True
     MsgBox "Ожидалось целое число", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_SrokFI_GE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_SrokFI_GE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_SrokFI_LE_Validate(cancel As Boolean)
If txtINVOS_INFO_SrokFI_LE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_SrokFI_LE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_SrokFI_LE.Text) <> CLng(val(txtINVOS_INFO_SrokFI_LE.Text)) Then
     cancel = True
     MsgBox "Ожидалось целое число", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_SrokFI_LE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_SrokFI_LE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_SrokOI_GE_Validate(cancel As Boolean)
If txtINVOS_INFO_SrokOI_GE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_SrokOI_GE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_SrokOI_GE.Text) <> CLng(val(txtINVOS_INFO_SrokOI_GE.Text)) Then
     cancel = True
     MsgBox "Ожидалось целое число", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_SrokOI_GE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_SrokOI_GE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_SrokOI_LE_Validate(cancel As Boolean)
If txtINVOS_INFO_SrokOI_LE.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtINVOS_INFO_SrokOI_LE.Text) Then
     cancel = True
     MsgBox "Ожидалось число", vbOKOnly + vbExclamation, "Внимание"
  ElseIf val(txtINVOS_INFO_SrokOI_LE.Text) <> CLng(val(txtINVOS_INFO_SrokOI_LE.Text)) Then
     cancel = True
     MsgBox "Ожидалось целое число", vbOKOnly + vbExclamation, "Внимание"
  End If
End If
End Sub
Private Sub txtINVOS_INFO_SrokOI_LE_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtINVOS_INFO_SrokOI_LE_Change()
  Changing
End Sub
Private Sub dtpINVOS_INFO_ActivateDate_GE_Change()
  Changing
End Sub
Private Sub dtpINVOS_INFO_ActivateDate_LE_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_Info_Change()
  Changing
End Sub
Private Sub txtINVOS_INFO_TechFilePath_Change()
  Changing
End Sub
Private Sub txtINVOS_PLACE_MatOtv_Change()
  Changing
End Sub
Private Sub cmdINVOS_PLACE_MatOtv_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_OWNER", id, brief) Then
          txtINVOS_PLACE_MatOtv.Tag = Left(id, 38)
          txtINVOS_PLACE_MatOtv = brief
        End If
End Sub
Private Sub txtINVOS_PLACE_TheHouse_Change()
  Changing
End Sub
Private Sub cmdINVOS_PLACE_TheHouse_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_BLD", id, brief) Then
          txtINVOS_PLACE_TheHouse.Tag = Left(id, 38)
          txtINVOS_PLACE_TheHouse = brief
        End If
End Sub
Private Sub txtINVOS_PLACE_ComplNumber_Change()
  Changing
End Sub
Private Sub txtINVOS_PLACE_DIrection_Change()
  Changing
End Sub
Private Sub cmdINVOS_PLACE_DIrection_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_DIR", id, brief) Then
          txtINVOS_PLACE_DIrection.Tag = Left(id, 38)
          txtINVOS_PLACE_DIrection = brief
        End If
End Sub
Private Sub txtINVOS_PLACE_Uprav_Change()
  Changing
End Sub
Private Sub cmdINVOS_PLACE_Uprav_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_UPR", id, brief) Then
          txtINVOS_PLACE_Uprav.Tag = Left(id, 38)
          txtINVOS_PLACE_Uprav = brief
        End If
End Sub
Private Sub txtINVOS_PLACE_Otdel_Change()
  Changing
End Sub
Private Sub cmdINVOS_PLACE_Otdel_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_OTDEL", id, brief) Then
          txtINVOS_PLACE_Otdel.Tag = Left(id, 38)
          txtINVOS_PLACE_Otdel = brief
        End If
End Sub
Private Sub txtINVOS_PLACE_Flow_Change()
  Changing
End Sub
Private Sub txtINVOS_PLACE_Room_Change()
  Changing
End Sub
Private Sub txtINVOS_PLACE_WorkPlaceNum_Change()
  Changing
End Sub
Private Sub txtINVOS_PLACE_TheOwner_Change()
  Changing
End Sub
Private Sub cmdINVOS_PLACE_TheOwner_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_OWNER", id, brief) Then
          txtINVOS_PLACE_TheOwner.Tag = Left(id, 38)
          txtINVOS_PLACE_TheOwner = brief
        End If
End Sub
Private Sub txtINVOS_PLACE_Info_Change()
  Changing
End Sub
Private Sub txtINVOS_CODE_VisibleCode_Change()
  Changing
End Sub
Private Sub dtpINVOS_SROK_RecalcDate_GE_Change()
  Changing
End Sub
Private Sub dtpINVOS_SROK_RecalcDate_LE_Change()
  Changing
End Sub
Private Sub PanelfGroupInit()
OnInit = True
Dim iii As Long ' for combo only

  txtINVOS_INFO_TheOrg.Tag = ""
  txtINVOS_INFO_TheOrg = ""
 LoadBtnPictures cmdINVOS_INFO_TheOrg, cmdINVOS_INFO_TheOrg.Tag
  cmdINVOS_INFO_TheOrg.RemoveAllMenu
cmbINVOS_INFO_IsMaterial.Clear
cmbINVOS_INFO_IsMaterial.AddItem "Да"
cmbINVOS_INFO_IsMaterial.ItemData(cmbINVOS_INFO_IsMaterial.NewIndex) = -1
cmbINVOS_INFO_IsMaterial.AddItem "Нет"
cmbINVOS_INFO_IsMaterial.ItemData(cmbINVOS_INFO_IsMaterial.NewIndex) = 0
  txtINVOS_INFO_OSType.Tag = ""
  txtINVOS_INFO_OSType = ""
 LoadBtnPictures cmdINVOS_INFO_OSType, cmdINVOS_INFO_OSType.Tag
  cmdINVOS_INFO_OSType.RemoveAllMenu
txtINVOS_INFO_Name = ""
txtINVOS_INFO_ShortName = ""
txtINVOS_INFO_INVNum = ""
txtINVOS_INFO_CardNum = ""
dtpINVOS_INFO_ActivateDate_GE = Date
dtpINVOS_INFO_ActivateDate_LE = Date
txtINVOS_INFO_TechFilePath = ""
  txtINVOS_PLACE_MatOtv.Tag = ""
  txtINVOS_PLACE_MatOtv = ""
 LoadBtnPictures cmdINVOS_PLACE_MatOtv, cmdINVOS_PLACE_MatOtv.Tag
  cmdINVOS_PLACE_MatOtv.RemoveAllMenu
  txtINVOS_PLACE_TheHouse.Tag = ""
  txtINVOS_PLACE_TheHouse = ""
 LoadBtnPictures cmdINVOS_PLACE_TheHouse, cmdINVOS_PLACE_TheHouse.Tag
  cmdINVOS_PLACE_TheHouse.RemoveAllMenu
txtINVOS_PLACE_ComplNumber = ""
  txtINVOS_PLACE_DIrection.Tag = ""
  txtINVOS_PLACE_DIrection = ""
 LoadBtnPictures cmdINVOS_PLACE_DIrection, cmdINVOS_PLACE_DIrection.Tag
  cmdINVOS_PLACE_DIrection.RemoveAllMenu
  txtINVOS_PLACE_Uprav.Tag = ""
  txtINVOS_PLACE_Uprav = ""
 LoadBtnPictures cmdINVOS_PLACE_Uprav, cmdINVOS_PLACE_Uprav.Tag
  cmdINVOS_PLACE_Uprav.RemoveAllMenu
  txtINVOS_PLACE_Otdel.Tag = ""
  txtINVOS_PLACE_Otdel = ""
 LoadBtnPictures cmdINVOS_PLACE_Otdel, cmdINVOS_PLACE_Otdel.Tag
  cmdINVOS_PLACE_Otdel.RemoveAllMenu
txtINVOS_PLACE_Flow = ""
txtINVOS_PLACE_Room = ""
txtINVOS_PLACE_WorkPlaceNum = ""
  txtINVOS_PLACE_TheOwner.Tag = ""
  txtINVOS_PLACE_TheOwner = ""
 LoadBtnPictures cmdINVOS_PLACE_TheOwner, cmdINVOS_PLACE_TheOwner.Tag
  cmdINVOS_PLACE_TheOwner.RemoveAllMenu
txtINVOS_CODE_VisibleCode = ""
dtpINVOS_SROK_RecalcDate_GE = Date
dtpINVOS_SROK_RecalcDate_LE = Date
OnInit = False
End Sub



