VERSION 5.00
Object = "{84E5CF37-E467-4AC2-89C4-C6002FFB5055}#25.1#0"; "ChartViewer.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DynChart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   10575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14910
   ScaleHeight     =   705
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   994
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   12240
      TabIndex        =   58
      Text            =   "Text2"
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdRewf 
      Height          =   375
      Left            =   1800
      Picture         =   "UserControl1.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Один час вперед"
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton cmdRewb 
      Height          =   375
      Left            =   840
      Picture         =   "UserControl1.ctx":038A
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Один час назад"
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   7320
      Picture         =   "UserControl1.ctx":0714
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Обновить"
      Top             =   8280
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.TextBox txtMinute 
      Height          =   315
      Left            =   14010
      TabIndex        =   52
      Text            =   "00"
      ToolTipText     =   "Выбор правого значения оси Х"
      Top             =   8325
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtHour 
      Height          =   315
      Left            =   13560
      TabIndex        =   51
      Text            =   "18"
      ToolTipText     =   "Выбор правого значения оси Х"
      Top             =   8325
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.ComboBox cmbHour 
      Height          =   315
      ItemData        =   "UserControl1.ctx":0A9E
      Left            =   9960
      List            =   "UserControl1.ctx":0AB3
      Style           =   2  'Dropdown List
      TabIndex        =   50
      ToolTipText     =   "Выбор длины диапазона оси X"
      Top             =   8325
      Width           =   1695
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   75
      TabIndex        =   45
      Text            =   "0"
      ToolTipText     =   "Ввод минимального значения в ручную"
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   120
      TabIndex        =   44
      Text            =   "0"
      ToolTipText     =   "Ввод максимального значения в ручную"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "OK"
      Height          =   255
      Left            =   3960
      TabIndex        =   43
      Top             =   11760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar blue 
      Height          =   255
      Left            =   3960
      Max             =   255
      TabIndex        =   42
      Top             =   11400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.HScrollBar green 
      Height          =   255
      Left            =   3960
      Max             =   255
      TabIndex        =   41
      Top             =   11040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.HScrollBar red 
      Height          =   255
      Left            =   3960
      Max             =   255
      TabIndex        =   40
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+10%"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   30
      Top             =   11640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar scbXDiapazon 
      Height          =   375
      Left            =   1920
      Max             =   600
      Min             =   10
      TabIndex        =   26
      Top             =   11520
      Value           =   60
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   7
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   9960
      Width           =   4620
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   6
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   9600
      Width           =   4620
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   5
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   9240
      Width           =   4620
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   4
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   8880
      Width           =   4620
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   9960
      Width           =   4620
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   9600
      Width           =   4620
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   9240
      Width           =   4620
   End
   Begin VB.ComboBox cmbAgrDiap 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "UserControl1.ctx":0AE2
      Left            =   6360
      List            =   "UserControl1.ctx":0AF5
      TabIndex        =   10
      Text            =   "3 min"
      ToolTipText     =   "Диапазон агрегации значений"
      Top             =   8325
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Выбор нагрузки для отображения"
      Top             =   8880
      Width           =   4620
   End
   Begin VB.TextBox txtTime 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Text            =   "18:00"
      Top             =   11640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   3000
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   315
      Left            =   11880
      TabIndex        =   5
      ToolTipText     =   "Выбор правого значения оси Х"
      Top             =   8325
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   56885249
      CurrentDate     =   39280
   End
   Begin CDChartViewer.ChartViewer ChartViewer1 
      Height          =   7695
      Left            =   840
      ToolTipText     =   "Поле с графиками"
      Top             =   360
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   13573
      BorderStyle     =   1
      AutoSize        =   0   'False
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-10%"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   31
      Top             =   11400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-10%"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   11160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+10%"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   0
      Top             =   10920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Варианты выборок:"
      Height          =   255
      Left            =   12960
      TabIndex        =   62
      Top             =   9120
      Width           =   1575
   End
   Begin MSForms.OptionButton cmdSelection 
      Height          =   375
      Index           =   2
      Left            =   14160
      TabIndex        =   61
      Top             =   9360
      Width           =   495
      VariousPropertyBits=   746588179
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "873;661"
      Value           =   "0"
      Caption         =   " 3"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton cmdSelection 
      Height          =   375
      Index           =   1
      Left            =   13560
      TabIndex        =   60
      Top             =   9360
      Width           =   495
      VariousPropertyBits=   746588179
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "873;661"
      Value           =   "0"
      Caption         =   "2"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton cmdSelection 
      Height          =   375
      Index           =   0
      Left            =   12960
      TabIndex        =   59
      Top             =   9360
      Width           =   495
      VariousPropertyBits=   746588179
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "873;661"
      Value           =   "0"
      Caption         =   "1"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblSeparat 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   13920
      TabIndex        =   53
      Top             =   8400
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSForms.ToggleButton YmaxLock 
      Height          =   375
      Left            =   195
      TabIndex        =   49
      ToolTipText     =   "Зафиксировать максимальное значение оси Y"
      Top             =   720
      Width           =   375
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "661;661"
      Value           =   "0"
      PicturePosition =   262148
      Picture         =   "UserControl1.ctx":0B18
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton YminLock 
      Height          =   375
      Left            =   150
      TabIndex        =   48
      ToolTipText     =   "Зафиксировать минимальное значение оси Y"
      Top             =   7080
      Width           =   375
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "661;661"
      Value           =   "0"
      PicturePosition =   262148
      Picture         =   "UserControl1.ctx":0EB2
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Y max:"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Y min:"
      Height          =   255
      Left            =   75
      TabIndex        =   46
      Top             =   7860
      Width           =   495
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   11400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   11040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   10680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4560
      Top             =   11760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   315
      Index           =   7
      Left            =   12240
      TabIndex        =   39
      Top             =   9960
      Width           =   315
      ForeColor       =   0
      BackColor       =   33023
      VariousPropertyBits=   25
      Size            =   "556;556"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   315
      Index           =   6
      Left            =   12240
      TabIndex        =   38
      Top             =   9600
      Width           =   315
      ForeColor       =   0
      BackColor       =   12632064
      VariousPropertyBits=   25
      Size            =   "556;556"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   315
      Index           =   5
      Left            =   12240
      TabIndex        =   37
      Top             =   9240
      Width           =   315
      ForeColor       =   0
      BackColor       =   0
      VariousPropertyBits=   25
      Size            =   "556;556"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   315
      Index           =   4
      Left            =   12240
      TabIndex        =   36
      Top             =   8880
      Width           =   315
      ForeColor       =   0
      BackColor       =   65535
      VariousPropertyBits=   25
      Size            =   "556;556"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   315
      Index           =   3
      Left            =   6000
      TabIndex        =   35
      Top             =   9960
      Width           =   315
      ForeColor       =   0
      BackColor       =   16711935
      VariousPropertyBits=   25
      Size            =   "556;556"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   315
      Index           =   2
      Left            =   6000
      TabIndex        =   34
      Top             =   9600
      Width           =   315
      ForeColor       =   0
      BackColor       =   16711680
      VariousPropertyBits=   25
      Size            =   "556;556"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   315
      Index           =   1
      Left            =   6000
      TabIndex        =   33
      Top             =   9240
      Width           =   315
      ForeColor       =   0
      BackColor       =   49152
      VariousPropertyBits=   25
      Size            =   "556;556"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line3 
      X1              =   848
      X2              =   848
      Y1              =   584
      Y2              =   696
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   315
      Index           =   0
      Left            =   6000
      TabIndex        =   32
      Top             =   8880
      Width           =   315
      ForeColor       =   0
      BackColor       =   255
      VariousPropertyBits=   25
      Size            =   "556;556"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblXdiap 
      BackStyle       =   0  'Transparent
      Caption         =   "min"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   29
      Top             =   11595
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblXdiap 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   28
      Top             =   11595
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblXdiap 
      BackStyle       =   0  'Transparent
      Caption         =   "h."
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   27
      Top             =   11640
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   25
      ToolTipText     =   "Показать/Скрыть график"
      Top             =   9960
      Width           =   255
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   23
      ToolTipText     =   "Показать/Скрыть график"
      Top             =   9600
      Width           =   255
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Index           =   5
      Left            =   7080
      TabIndex        =   21
      ToolTipText     =   "Показать/Скрыть график"
      Top             =   9240
      Width           =   255
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   19
      ToolTipText     =   "Показать/Скрыть график"
      Top             =   8880
      Width           =   255
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   17
      ToolTipText     =   "Показать/Скрыть график"
      Top             =   9960
      Width           =   255
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   15
      ToolTipText     =   "Показать/Скрыть график"
      Top             =   9600
      Width           =   255
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   13
      ToolTipText     =   "Показать/Скрыть график"
      Top             =   9240
      Width           =   255
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   11
      ToolTipText     =   "Показать/Скрыть график"
      Top             =   8880
      Width           =   255
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line6 
      X1              =   -8
      X2              =   984
      Y1              =   584
      Y2              =   584
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "18:00:00"
      Height          =   255
      Left            =   13560
      TabIndex        =   8
      Top             =   8325
      Width           =   735
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "17.07.2007"
      Height          =   255
      Left            =   12240
      TabIndex        =   7
      Top             =   8325
      Width           =   975
   End
   Begin MSForms.ToggleButton cmdAgregate 
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Усреднить значения"
      Top             =   8295
      Width           =   375
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "661;661"
      Value           =   "0"
      PicturePosition =   262148
      Picture         =   "UserControl1.ctx":124C
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton cmdPlay 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Реальное время/Архив"
      Top             =   8280
      Width           =   375
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "661;661"
      Value           =   "1"
      PicturePosition =   262148
      Picture         =   "UserControl1.ctx":15E6
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line4 
      X1              =   528
      X2              =   528
      Y1              =   548
      Y2              =   584
   End
   Begin VB.Line Line2 
      X1              =   48
      X2              =   48
      Y1              =   0
      Y2              =   696
   End
   Begin VB.Line Line1 
      X1              =   -8
      X2              =   984
      Y1              =   548
      Y2              =   548
   End
   Begin VB.Label lblXdiap 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   11595
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   10455
      Left            =   0
      Top             =   0
      Width           =   14775
   End
End
Attribute VB_Name = "DynChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private Sub blue_Change()
Shape2.FillColor = RGB(red.Value, green.Value, blue.Value)

End Sub

Private Sub CheckBox1_Click(Index As Integer)
'If Combo1(Index).ItemData(Combo1(Index).ListIndex) > 9990 And CheckBox1(Index).Value = True Then
'ag = cmdAgregate.Value
'ad = cmbAgrDiap.ListIndex
'cmdAgregate.Value = True
'cmbAgrDiap.ListIndex = 3
'Else
'cmdAgregate.Value = ag
'cmbAgrDiap.ListIndex = ad
'End If
Call Refresh
End Sub

Private Sub max_value_chk(ByVal d1 As String, ByVal d2 As String)
flag = 0
For i = 0 To 7
If CheckBox1(i).Value = True And Combo1(i).ListIndex >= 0 Then
f = 1
Dim mdbo As New ADODB.Connection
mdbo.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=nb;Initial Catalog=OPINFO;Data Source=SQL"
mdbo.Open
Dim mdboc As New ADODB.Command
mdboc.ActiveConnection = mdbo
Dim mdbor As New ADODB.Recordset
If Combo1(i).ItemData(Combo1(i).ListIndex) > 9990 And CheckBox1(i).Value = True Then
mdboc.CommandText = "Select * from All_obj where id=" & Combo1(i).ItemData(Combo1(i).ListIndex)
mdbor.Open mdboc
afield = mdbor(2).Value
atable = mdbor(3).Value
mdbor.Close

'd1 = Year(d1) & "-" & Month(d1) & "-" & Day(d1) & " " & Hour(d1) & ":00:00.000"
'd2 = Year(d2) & "-" & Month(d2) & "-" & Day(d2) & " " & Hour(d2) & ":00:00.000"
'dat = "CAST((CAST(DATEPART(yyyy, datee) as varchar) + '-' + CAST(month(datee) as varchar)+ '-' + CAST(day(datee) as varchar) + ' ' + left(timee,2) + ':00:00.000') as datetime)"
Y = Year(d2)
m = Month(d2)
d = Day(d2)
h = Hour(d2)
h = h - 1
If h < 0 Then
h = 24 + h
d = d - 1
End If
If d < 1 Then
m = m - 1
d = dayy(m)
End If
If m < 1 Then
m = 12
Y = Y - 1
d = dayy(m)
End If
d2m = Y & "-" & m & "-" & d & " " & h & ":00:00.000"

dat = "datez"
mdboc.CommandText = "Select MAX(" & afield & ") from " & atable & " where " & dat & "<='" & d1 & "' and " & dat & ">'" & d2m & "' "
mdbor.Open mdboc
mv = mdbor(0).Value
If mv < mav Then
Else
mav = mv
End If

If mdbor(0).Value & "e" = "e" Then
'flag = 1
CheckBox1(i).Value = False
mdbor.Close
Exit Sub
End If
mdbor.Close

mdboc.CommandText = "Select MIN(" & afield & ") from " & atable & " where " & dat & "<='" & d1 & "' and " & dat & ">'" & d2m & "' "
mdbor.Open mdboc
If mdbor(0).Value = 0 Then
mv = 1
Else
mv = mdbor(0).Value
End If


If mv > miv And miv <> Empty Then
Else
miv = mv
End If
mdbor.Close
mdbo.Close

Else
mdboc.CommandText = "Select * from All_obj where id=" & Combo1(i).ItemData(Combo1(i).ListIndex)
mdbor.Open mdboc
afield = mdbor(2).Value
atable = mdbor(3).Value
mdbor.Close
'2007-07-20 09:35:00.000
mdboc.CommandText = "Select MAX(" & afield & ") from " & atable & " where data<='" & d1 & "' and data>='" & d2 & "'"
mdbor.Open mdboc
mv = mdbor(0).Value + 1
If mv < mav Then
Else
mav = mv
End If

If mdbor(0).Value & "e" = "e" Then
mdbor.Close
'flag = 1
CheckBox1(i).Value = False
Exit Sub
End If
mdbor.Close

mdboc.CommandText = "Select MIN(" & afield & ") from " & atable & " where data<='" & d1 & "' and data>='" & d2 & "'"
mdbor.Open mdboc
If mdbor(0).Value = 0 Then
mv = 1
Else
mv = mdbor(0).Value
End If


If mv > miv And miv <> Empty Then
Else
miv = mv
End If
mdbor.Close
mdbo.Close
End If
End If
Next i
If YmaxLock.Value = True Then
Else
tmav = CInt(mav)
txtMax.Text = Val(tmav)
End If
If YminLock.Value = True Then
Else
tmiv = CInt(miv)
txtMin.Text = Val(tmiv)
End If
End Sub


Private Sub cmbAgrDiap_Click()
If cmbAgrDiap.ItemData(cmbAgrDiap.ListIndex) * 2 >= cmbHour.ItemData(cmbHour.ListIndex) * 60 Then
cmbAgrDiap.ListIndex = 1
End If
End Sub

Private Sub cmbHour_Click()
Call Refresh

End Sub

Private Sub cmdAgregate_Click()
Call Refresh
End Sub

Private Sub cmdPlay_Click()
If cmdPlay.Value = True Then
'txtTime.Visible = False
txtHour.Visible = False
txtMinute.Visible = False
lblSeparat.Visible = False
txtDate.Visible = False
lblTime.Visible = True
lblDate.Visible = True
Call Refresh
Timer1.Enabled = True
Else
'txtTime.Visible = True
txtHour.Visible = True
txtMinute.Visible = True
lblSeparat.Visible = True

txtDate.Visible = True
lblTime.Visible = False
lblDate.Visible = False
Timer1.Enabled = False

txtDate.Value = Date
txtTime.Text = "00:00"
txtMinute.Text = "00"
txtHour.Text = "00"
End If
End Sub

Private Function dayy(ByVal m As Byte) As Byte
Select Case m
Case 1
dayy = 31
Case 2
dayy = 30
Case 3
dayy = 31
Case 4
dayy = 30
Case 5
dayy = 31
Case 6
dayy = 30
Case 7
dayy = 31
Case 8
dayy = 31
Case 9
dayy = 30
Case 10
dayy = 31
Case 11
dayy = 30
Case 12
dayy = 31
End Select
End Function

Private Sub draw_grph(ByVal d1 As String, ByVal d2 As String)
Dim kr As Single
Dim data0() As Integer
Dim data1() As Integer
Dim data2() As Integer
Dim data3() As Integer
Dim data4() As Integer
Dim data5() As Integer
Dim data6() As Integer
Dim data7() As Integer
Dim mc(7) As Integer
Dim f3(7) As Integer
Dim labels() As String
If cmdAgregate.Value = True Then
l = cmbAgrDiap.ItemData(cmbAgrDiap.ListIndex)
Else
l = 1
End If

lblDate.Caption = Date
lblTime.Caption = Hour(Time) & ":" & strdup("0", 2 - Len(Minute(Time))) & Minute(Time)
Dim data(7, 1440) As Integer
f = 0
f2 = 0
Dim mdbo As New ADODB.Connection
mdbo.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=nb;Initial Catalog=OPINFO;Data Source=SQL"
mdbo.Open
Dim mdboc As New ADODB.Command
mdboc.ActiveConnection = mdbo
Dim mdbor As New ADODB.Recordset



For i = 0 To 7
If CheckBox1(i).Value = True And Combo1(i).ListIndex >= 0 Then
If Combo1(i).ItemData(Combo1(i).ListIndex) > 9990 And CheckBox1(i).Value = True Then
f2 = 2
f3(i) = 2
f = 1
mdboc.CommandText = "Select * from All_obj where id=" & Combo1(i).ItemData(Combo1(i).ListIndex)
mdbor.Open mdboc
afield = mdbor(2).Value
atable = mdbor(3).Value
mdbor.Close
'ReDim labels(cmbHour.ItemData(cmbHour.ListIndex) - 1)
'd1m = Year(d1) & "-" & Month(d1) & "-" & Day(d1) & " " & Hour(d1) + 1 & ":00:00.000"
Y = Year(d2)
m = Month(d2)
d = Day(d2)
h = Hour(d2)
h = h - 1
If h < 0 Then
h = 24 + h
d = d - 1
End If
If d < 1 Then
m = m - 1
d = dayy(m)
End If
If m < 1 Then
m = 12
Y = Y - 1
d = dayy(m)
End If
d2m = Y & "-" & m & "-" & d & " " & h & ":00:00.000"
'dat = "CAST((CAST(DATEPART(yyyy, datee) as varchar) + '-' + CAST(month(datee) as varchar)+ '-' + CAST(day(datee) as varchar) + ' ' + left(timee,2) + ':00:00.000') as datetime)"
dat = "datez"
mdboc.CommandText = "Select distinct " & afield & ",timee,datez from " & atable & " where " & dat & "<='" & d1 & "' and " & dat & ">'" & d2m & "' order by datez asc"
Text1.Text = mdboc.CommandText
mdbor.Open mdboc
    j = 0
    jg = 0
    st = 1
Do Until mdbor.EOF
If jg = 0 Then
st = (Minute(d1) \ l) + 1
Else
If Val(Left(mdbor(1).Value, 2)) = Val(Hour(d1)) Then
st = (60 - Minute(d1)) \ l + 1
Else
st = 1
End If
End If
For kw = st To 60 \ l
data(i, j) = mdbor(0).Value
' labels(j) = mdbor(1).Value
j = j + 1
Next kw
jg = jg + 1
mdbor.MoveNext
Loop
mc(i) = j

mdbor.Close


Else
f = 1
mdboc.CommandText = "Select * from All_obj where id=" & Combo1(i).ItemData(Combo1(i).ListIndex)
mdbor.Open mdboc
afield = mdbor(2).Value
atable = mdbor(3).Value
mdbor.Close
'2007-07-20 09:35:00.000
mdboc.CommandText = "Select " & afield & " from " & atable & " where data<='" & d1 & "' and data>'" & d2 & "' order by data asc"
Text1.Text = mdboc.CommandText
mdbor.Open mdboc
    j = 0
Do Until mdbor.EOF

If cmdAgregate.Value = True Then
li = li + 1
ls = ls + mdbor(0).Value
If li = l Then
ls = ls \ l
data(i, j) = ls
j = j + 1
ls = 0
li = 0
End If
Else
data(i, j) = mdbor(0).Value
j = j + 1
End If
mdbor.MoveNext

Loop
mc(i) = j
mdbor.Close

End If
Else
mc(i) = 1
End If

Next i
mdbo.Close
If f = 1 Then
a = mc(0)
ReDim data0(a - 1) As Integer
a = mc(1)
ReDim data1(a - 1) As Integer
a = mc(2)
ReDim data2(a - 1) As Integer
a = mc(3)
ReDim data3(a - 1) As Integer
a = mc(4)
ReDim data4(a - 1) As Integer
a = mc(5)
ReDim data5(a - 1) As Integer
a = mc(6)
ReDim data6(a - 1) As Integer
a = mc(7)
ReDim data7(a - 1) As Integer

For k = 0 To 1440
If mc(0) > k Then
data0(k) = data(0, k)
End If
If mc(1) > k Then
data1(k) = data(1, k)
End If
If mc(2) > k Then
data2(k) = data(2, k)
End If
If mc(3) > k Then
data3(k) = data(3, k)
End If
If mc(4) > k Then
data4(k) = data(4, k)
End If
If mc(5) > k Then
data5(k) = data(5, k)
End If
If mc(6) > k Then
data6(k) = data(6, k)
End If
If mc(7) > k Then
data7(k) = data(7, k)
End If
Next k
If f2 = 3 Then
Else



a1 = cmbHour.ItemData(cmbHour.ListIndex) * 60 '\ l
et = l
a = a1 \ l
ai = 1
If a1 \ et > 60 Then
et = cmbHour.ItemData(cmbHour.ListIndex)
kr = (a1 \ l) / 60
a = a1 \ l
ai = kr
End If

ReDim labels(a - 1)
mi = Minute(d1)
h = Hour(d1)
d = Day(d1)
m = Month(d1)
Y = Year(d1)
dt = ""
dt = Y & "-" & m & "-" & d & " "

For i = 1 To a1 Step et
If i > a1 - cmbHour.ItemData(cmbHour.ListIndex) Then
dt = Y & "-" & m & "-" & d & " "
Else
End If
If cmdAgregate.Value = True Then
labels(a - 1) = dt & h & ":" & mi
Else
labels(a1 - i) = dt & h & ":" & mi
End If

h = h - et \ 60
b = et Mod 60
mi = mi - b
dt = ""
fl = 0
If mi < 0 Then
mi = 60 + mi
h = h - 1
End If
If h < 0 Then
h = 24 + h
d = d - 1
fl = 1
End If
If d = 0 Then
m = m - 1
fl = 1
d = dayy(m)
End If
If m = 0 Then
m = 12
Y = Y - 1
fl = 1
d = dayy(m)
End If
If fl = 1 Then
dt = Y & "-" & m & "-" & d & " "

End If
a = a - ai
Next i
End If

Dim cd As New ChartDirector.API
Dim c As XYChart
Set c = cd.XYChart(900, 500, &HFFFFFF, &H0, 0)
Call c.setPlotArea(50, 30, 800, 370, &HFFFFFF, -1, -1, &HCCCCCC, &HCCCCCC)
Call c.yAxis().setTitle("MWatts")
Call c.yAxis().setLinearScale(tmiv - ll, tmav + ul)
Call c.xAxis().setLabels(labels).setFontAngle(90)
Call c.xAxis().setLabels(labels).setFontSize(8)
Call c.yAxis2().setTitle("")
Call c.yAxis2().setLinearScale(tmiv - ll, tmav + ul)
For i = 0 To 7
mrk = data(i, mc(i) - 1)
 a = strdup("0", 6 - Len(Hex(CommandButton1(i).BackColor))) & Hex(CommandButton1(i).BackColor)
 
Call c.yAxis2().addMark(mrk, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Str(mrk)).setLineWidth(0)
Next i

Dim layer As LineLayer
Dim layer0 As StepLineLayer
Set layer0 = c.addStepLineLayer()
Call layer0.setLineWidth(2)
Set layer = c.addLineLayer2()
Call layer.setLineWidth(2)
If f3(0) = 0 Then
    a = strdup("0", 6 - Len(Hex(CommandButton1(0).BackColor))) & Hex(CommandButton1(0).BackColor)
    Call layer.addDataSet(data0, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(0).List(Combo1(0).ListIndex))
Else
    a = strdup("0", 6 - Len(Hex(CommandButton1(0).BackColor))) & Hex(CommandButton1(0).BackColor)
    Call layer0.addDataSet(data0, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(0).List(Combo1(0).ListIndex))
End If
If f3(1) = 0 Then
    a = strdup("0", 6 - Len(Hex(CommandButton1(1).BackColor))) & Hex(CommandButton1(1).BackColor)
    Call layer.addDataSet(data1, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(1).List(Combo1(1).ListIndex))
Else
    a = strdup("0", 6 - Len(Hex(CommandButton1(1).BackColor))) & Hex(CommandButton1(1).BackColor)
    Call layer0.addDataSet(data1, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(1).List(Combo1(1).ListIndex))
End If
If f3(2) = 0 Then
    a = strdup("0", 6 - Len(Hex(CommandButton1(2).BackColor))) & Hex(CommandButton1(2).BackColor)
    Call layer.addDataSet(data2, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(2).List(Combo1(2).ListIndex))
Else
    a = strdup("0", 6 - Len(Hex(CommandButton1(2).BackColor))) & Hex(CommandButton1(2).BackColor)
    Call layer0.addDataSet(data2, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(2).List(Combo1(2).ListIndex))
End If
If f3(3) = 0 Then
    a = strdup("0", 6 - Len(Hex(CommandButton1(3).BackColor))) & Hex(CommandButton1(3).BackColor)
    Call layer.addDataSet(data3, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(3).List(Combo1(3).ListIndex))
Else
    a = strdup("0", 6 - Len(Hex(CommandButton1(3).BackColor))) & Hex(CommandButton1(3).BackColor)
    Call layer0.addDataSet(data3, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(3).List(Combo1(3).ListIndex))
End If
If f3(4) = 0 Then
    a = strdup("0", 6 - Len(Hex(CommandButton1(4).BackColor))) & Hex(CommandButton1(4).BackColor)
    Call layer.addDataSet(data4, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(4).List(Combo1(4).ListIndex))
Else
    a = strdup("0", 6 - Len(Hex(CommandButton1(4).BackColor))) & Hex(CommandButton1(4).BackColor)
    Call layer0.addDataSet(data4, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(4).List(Combo1(4).ListIndex))
End If
If f3(5) = 0 Then
    a = strdup("0", 6 - Len(Hex(CommandButton1(5).BackColor))) & Hex(CommandButton1(5).BackColor)
    Call layer.addDataSet(data5, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(5).List(Combo1(5).ListIndex))
Else
    a = strdup("0", 6 - Len(Hex(CommandButton1(5).BackColor))) & Hex(CommandButton1(5).BackColor)
    Call layer0.addDataSet(data5, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(5).List(Combo1(5).ListIndex))
End If
If f3(6) = 0 Then
    a = strdup("0", 6 - Len(Hex(CommandButton1(6).BackColor))) & Hex(CommandButton1(6).BackColor)
    Call layer.addDataSet(data6, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(6).List(Combo1(6).ListIndex))
Else
    a = strdup("0", 6 - Len(Hex(CommandButton1(6).BackColor))) & Hex(CommandButton1(6).BackColor)
    Call layer0.addDataSet(data6, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(6).List(Combo1(6).ListIndex))
End If
If f3(7) = 0 Then
    a = strdup("0", 6 - Len(Hex(CommandButton1(7).BackColor))) & Hex(CommandButton1(7).BackColor)
    Call layer.addDataSet(data7, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(7).List(Combo1(7).ListIndex))
Else
    a = strdup("0", 6 - Len(Hex(CommandButton1(7).BackColor))) & Hex(CommandButton1(7).BackColor)
    Call layer0.addDataSet(data7, "&H" & Mid(a, 5, 2) & Mid(a, 3, 2) & Mid(a, 1, 2), Combo1(7).List(Combo1(7).ListIndex))
End If
        Set ChartViewer1.Picture = c.makePicture()
Else
    Set ChartViewer1.Picture = Nothing
End If


End Sub



Private Sub cmdRefresh_Click()
Text1.Text = b
'If Combo1(Index).ItemData(Combo1(Index).ListIndex) > 9990 And CheckBox1(Index).Value = True Then
'ag = cmdAgregate.Value
'ad = cmbAgrDiap.ListIndex
'cmdAgregate.Value = True
'cmbAgrDiap.ListIndex = 3
'Else
'cmdAgregate.Value = ag
'cmbAgrDiap.ListIndex = ad
'End If
Call Refresh
End Sub

Private Sub cmdRewb_Click()
txtHour.Text = Val(txtHour.Text) - 1
End Sub

Private Sub cmdRewf_Click()
txtHour.Text = Val(txtHour.Text) + 1
End Sub

Private Sub cmdSelection_Change(Index As Integer)
If cmdSelection(Index).Value = True Then
Dim mdbo As New ADODB.Connection
login = Environ("USERNAME")

mdbo.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=nb;Initial Catalog=OPINFO;Data Source=SQL"
mdbo.Open
Dim mdboc As New ADODB.Command
mdboc.ActiveConnection = mdbo
Dim mdbor As New ADODB.Recordset
Dim a(7) As Integer

If prsel = 0 Then
Else
For i = 0 To 7
If CheckBox1(i).Value = True Then
a(i) = Combo1(i).ListIndex
Else
a(i) = -1
End If
Next i
mdboc.CommandText = "update dc_Selections set field1='" & a(0) & "', field2='" & a(1) & "', field3='" & a(2) & "', field4='" & a(3) & "', field5='" & a(4) & "', field6='" & a(5) & "', field7='" & a(6) & "', field8='" & a(7) & "' WHERE login='" & login & "' and Selnum=" & prsel - 1
mdbor.Open mdboc
End If

mdboc.CommandText = "Select * from dc_Selections WHERE login ='" & login & "' and SelNum='" & Index & "'"
mdbor.Open mdboc
If mdbor.EOF = True Then
mdbor.Close
mdboc.CommandText = "insert into dc_Selections (SelNum, Login) Values (" & Index & ", '" & login & "')"
mdbor.Open mdboc
For i = 0 To 7
Combo1(i).ListIndex = -1
CheckBox1(i).Value = False
Next i

Else
For i = 0 To 7
Combo1(i).ListIndex = mdbor(i + 3).Value
If mdbor(i + 3).Value <> "-1" Then
CheckBox1(i).Value = True
Else
CheckBox1(i).Value = False
End If
Next i
mdbor.Close
End If

mdboc.CommandText = "update dc_LastChoice Set lastselnum = " & Index & " Where login='" & login & "'"
mdbor.Open mdboc
prsel = Index + 1
End If
End Sub

Private Sub Combo1_Click(Index As Integer)
'If Combo1(Index).ItemData(Combo1(Index).ListIndex) > 9990 And CheckBox1(Index).Value = True Then
'ag = cmdAgregate.Value
'ad = cmbAgrDiap.ListIndex
'cmdAgregate.Value = True
'cmbAgrDiap.ListIndex = 3
'Else
'cmdAgregate.Value = ag
'cmbAgrDiap.ListIndex = ad
'End If
Call Refresh
End Sub

Private Sub Command1_Click()
ul = ul + ((tmav - tmiv) * 0.1)
End Sub

Private Sub Command2_Click()
ul = ul - ((tmav - tmiv) * 0.1)
End Sub

Private Sub Command3_Click()
ll = ll + ((tmav - tmiv) * 0.1)
End Sub

Private Sub Command4_Click()
ll = ll - ((tmav - tmiv) * 0.1)
End Sub

Private Sub Command5_Click()
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
red.Visible = False
green.Visible = False
blue.Visible = False
Command5.Visible = False
CommandButton1(ii).BackColor = RGB(red.Value, green.Value, blue.Value)
End Sub




Private Sub CommandButton1_Click(Index As Integer)
Shape2.Visible = True
Shape3.Visible = True
Shape4.Visible = True
Shape5.Visible = True
red.Visible = True
green.Visible = True
blue.Visible = True
Command5.Visible = True
Shape2.FillColor = CommandButton1(Index).BackColor
a = strdup("0", 6 - Len(Hex(CommandButton1(Index).BackColor))) & Hex(CommandButton1(Index).BackColor)
'Label1.Caption = Val("&H" & Mid(a, 9, 2))
red.Value = Val("&H" & Mid(a, 5, 2))
green.Value = Val("&H" & Mid(a, 3, 2))
blue.Value = Val("&H" & Mid(a, 1, 2))
ii = Index
End Sub
Private Function strdup(ByVal s As String, ByVal n As Integer) As String
For i = 1 To n
z = z & s
Next i
strdup = z
End Function

Private Sub green_Change()
Shape2.FillColor = RGB(red.Value, green.Value, blue.Value)
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub red_Change()
Shape2.FillColor = RGB(red.Value, green.Value, blue.Value)
End Sub

Private Sub scbTime_Change()
txtTime.Visible = True
txtDate.Visible = True
lblTime.Visible = False
lblDate.Visible = False
Timer1.Enabled = False
End Sub

Private Sub scbXDiapazon_Change()
If scbXDiapazon.Value < 60 Then
lblXdiap(2).Caption = scbXDiapazon.Value
lblXdiap(3).Caption = "min"

Else
a = scbXDiapazon.Value \ 60
lblXdiap(0).Caption = a
lblXdiap(1).Caption = "h."
lblXdiap(2).Caption = scbXDiapazon.Value - (a * 60)
lblXdiap(3).Caption = "min"
End If
End Sub

Private Sub Text2_Change()
If Text2.Text = "-1" Then
For i = 0 To 7
CheckBox1(i).Visible = True
Combo1(i).Visible = True
CommandButton1(i).Visible = True
UserControl.Height = 10605
Next i
Else
UserControl.Height = 8790
For i = 0 To 7
CheckBox1(i).Visible = False
Combo1(i).Visible = False
CommandButton1(i).Visible = False
Next i
Combo1(0).ListIndex = CInt(Val(Text2.Text))
CheckBox1(0).Value = True
End If
End Sub

Private Sub Timer1_Timer()
'Text1.Text = bloca
d1 = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " " & Hour(Time) & ":" & Minute(Time)
Y = Year(Date)
m = Month(Date)
d = Day(Date)
h = Hour(Time) - cmbHour.ItemData(cmbHour.ListIndex) '- CInt(lblXdiap(0))
mi = Minute(Time) - CInt(lblXdiap(2))
If mi < 0 Then
mi = 60 + mi
h = h - 1
End If
If h < 0 Then
h = 24 + h
d = d - 1
End If
If d < 1 Then
m = m - 1
d = dayy(m)
End If
If m < 1 Then
m = 12
Y = Y - 1
d = dayy(m)
End If

d2 = Y & "-" & m & "-" & d & " " & h & ":" & mi
Call draw_grph(d1, d2)
End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub txtDate_Change()
Call Refresh
End Sub

Private Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."

If cmdPlay.Value = True Then
d1 = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " " & Hour(Time) & ":" & Minute(Time)
Y = Year(Date)
m = Month(Date)
d = Day(Date)
h = Hour(Time) - cmbHour.ItemData(cmbHour.ListIndex)
mi = Minute(Time) - CInt(lblXdiap(2))
If mi < 0 Then
mi = 60 + mi
h = h - 1
End If
If h < 0 Then
h = 24 + h
d = d - 1
End If
If d < 1 Then
m = m - 1
d = dayy(m)
End If
If m < 1 Then
m = 12
Y = Y - 1
d = dayy(m)
End If
d2 = Y & "-" & strdup(0, 2 - Len(m)) & m & "-" & strdup(0, 2 - Len(d)) & d & " " & strdup(0, 2 - Len(h)) & h & ":" & mi
Else
d1 = Year(txtDate.Value) & "-" & Month(txtDate.Value) & "-" & Day(txtDate.Value) & " " & Hour(txtTime.Text) & ":" & Minute(txtTime.Text)
Y = Year(txtDate.Value)
m = Month(txtDate.Value)
d = Day(txtDate.Value)
h = Hour(txtTime.Text) - cmbHour.ItemData(cmbHour.ListIndex) '- CInt(lblXdiap(0))
mi = Minute(txtTime.Text) - CInt(lblXdiap(2))
If mi < 0 Then
mi = 60 + mi
h = h - 1
End If
If h < 0 Then
h = 24 + h
d = d - 1
End If
If d < 1 Then
m = m - 1
d = dayy(m)
End If
If m < 1 Then
m = 12
Y = Y - 1
d = dayy(m)
End If
d2 = Y & "-" & strdup(0, 2 - Len(m)) & m & "-" & strdup(0, 2 - Len(d)) & d & " " & strdup(0, 2 - Len(h)) & h & ":" & mi
End If

Call max_value_chk(d1, d2)
If flag = 1 Then
cmdPlay.Value = True
Else
Call draw_grph(d1, d2)
End If
End Sub
Private Sub txtHour_Change()
a = Val(Left(txtHour.Text, 2))
If a >= 24 Then
a = a - 24
d = Day(txtDate.Value)
m = Month(txtDate.Value)
Y = Year(txtDate.Value)
d = d + 1
If d > dayy(m) Then
m = m + 1
d = 1
End If
If m > 12 Then
m = 1
Y = Y + 1
End If
txtDate.Value = d & "." & m & "." & Y
End If
If a < 0 Then
d = Day(txtDate.Value)
m = Month(txtDate.Value)
Y = Year(txtDate.Value)
a = 24 + a
d = d - 1
If d < 1 Then
m = m - 1
d = dayy(m)
End If
If m < 1 Then
m = 12
Y = Y - 1
d = dayy(m)
End If
txtDate.Value = d & "." & m & "." & Y
End If
txtHour.Text = a
txtTime.Text = strdup("0", 2 - Len(a)) & a & ":" & txtMinute.Text
End Sub

Private Sub txtMax_Change()
tmav = CInt(Val(txtMax.Text))
End Sub

Private Sub txtMin_Change()
tmiv = CInt(Val(txtMin.Text))
End Sub

Private Sub txtMinute_Change()
a = Val(Left(txtMinute.Text, 2))
If a >= 60 Then
a = 0
End If
If a < 0 Then
a = 60 + a
txtHour.Text = Int(Val(txtHour.Text)) - 1
End If

txtMinute.Text = strdup("0", 2 - Len(a)) & a

txtTime.Text = txtHour.Text & ":" & txtMinute.Text
End Sub

Private Sub txtTime_Change()
Call Refresh
End Sub

Private Sub UserControl_Initialize()
'Checking the security and autorization
Dim mdbo As New ADODB.Connection
mdbo.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=nb;Initial Catalog=NASTJA;Data Source=SQL"
mdbo.Open
Dim mdboc As New ADODB.Command
mdboc.ActiveConnection = mdbo
Dim mdbor As New ADODB.Recordset
mdboc.CommandText = "Select DynChart from dostup where kasutajanimi='" & Environ("USERNAME") & "'"
mdbor.Open mdboc
dostup = mdbor(0).Value
mdbor.Close
mdbo.Close

'Loading the list of objects
cmbHour.ListIndex = 0
cmbAgrDiap.ListIndex = 0
mdbo.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=nb;Initial Catalog=OPINFO;Data Source=SQL"
mdbo.Open
mdboc.ActiveConnection = mdbo
mdboc.CommandText = "Select * from All_obj WHERE " & dostup
mdbor.Open mdboc
Do Until mdbor.EOF
For i = 0 To 7
Combo1(i).AddItem (mdbor(1).Value)
Combo1(i).ItemData(j) = mdbor(0).Value
Next i
mdbor.MoveNext
j = j + 1
Loop
mdbor.Close
'Loading last selection
login = Environ("USERNAME")
mdboc.CommandText = "Select * from dc_LastChoice WHERE login ='" & login & "'"
mdbor.Open mdboc
If mdbor.EOF = True Then
mdbor.Close
mdboc.CommandText = "Insert into dc_LastChoice (login, lastselnum) values ('" & login & "',0)"
mdbor.Open mdboc
cmdSelection(0).Value = True
Else
cmdSelection(mdbor(2).Value).Value = True
End If

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text2,Text2,-1,Text
Public Property Get block() As String
Attribute block.VB_Description = "Returns/sets the text contained in the control."
    block = Text2.Text
End Property

Public Property Let block(ByVal New_block As String)
    Text2.Text() = New_block
    PropertyChanged "block"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Text2.Text = PropBag.ReadProperty("block", "-1")
End Sub

Private Sub UserControl_Terminate()
Dim mdbo As New ADODB.Connection
login = Environ("USERNAME")

mdbo.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=nb;Initial Catalog=OPINFO;Data Source=SQL"
mdbo.Open
Dim mdboc As New ADODB.Command
mdboc.ActiveConnection = mdbo
Dim mdbor As New ADODB.Recordset

If prsel = 0 Then
Else
mdboc.CommandText = "update dc_Selections set field1='" & Combo1(0).ListIndex & "', field2='" & Combo1(1).ListIndex & "', field3='" & Combo1(2).ListIndex & "', field4='" & Combo1(3).ListIndex & "', field5='" & Combo1(4).ListIndex & "', field6='" & Combo1(5).ListIndex & "', field7='" & Combo1(6).ListIndex & "', field8='" & Combo1(7).ListIndex & "' WHERE login='" & login & "' and Selnum=" & prsel - 1
mdbor.Open mdboc
End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("block", Text2.Text, "-1")
End Sub

