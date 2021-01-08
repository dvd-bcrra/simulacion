VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vbForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proyecto de simulacion"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14700
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   14700
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox todo_list 
      Height          =   7185
      ItemData        =   "Proyecto.frx":0000
      Left            =   0
      List            =   "Proyecto.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   3585
   End
   Begin VB.Frame Panel 
      Caption         =   "Aplicación - Conclusión"
      Height          =   7215
      Index           =   9
      Left            =   3720
      TabIndex        =   154
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   9360
         TabIndex        =   156
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   155
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label lblFrecuencia 
         Caption         =   "sin datos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   167
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label72 
         Caption         =   "Número de muestras del contaminante:"
         Height          =   255
         Left            =   2280
         TabIndex        =   166
         Top             =   4680
         Width           =   3735
      End
      Begin VB.Label lblContaminante 
         Caption         =   "sin datos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   165
         Top             =   4200
         Width           =   4815
      End
      Begin VB.Label Label71 
         Caption         =   "Contaminante con mayor frecuencia: "
         Height          =   255
         Left            =   2280
         TabIndex        =   164
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label lblApta 
         Caption         =   "sin datos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   163
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label69 
         Caption         =   "El agua es apta para la vida:"
         Height          =   255
         Left            =   2280
         TabIndex        =   162
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label67 
         Caption         =   "Por lo tanto"
         Height          =   255
         Left            =   4440
         TabIndex        =   161
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblPredomina 
         Caption         =   "sin datos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   160
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label68 
         Caption         =   "Predomina:"
         Height          =   255
         Left            =   2280
         TabIndex        =   159
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label lblPuntosAb 
         Caption         =   "sin datos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   158
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label66 
         Caption         =   "Puntos del abrevadero donde el agua es apta para la vida:"
         Height          =   255
         Left            =   2280
         TabIndex        =   157
         Top             =   720
         Width           =   5175
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Generar numeros pseudoaleatorios"
      Height          =   7215
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton btn_CalcularN 
         Caption         =   "Calcular"
         Height          =   375
         Left            =   5520
         TabIndex        =   84
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton bnt_GenerarDefault 
         Caption         =   "Generar default"
         Height          =   495
         Left            =   360
         TabIndex        =   80
         Top             =   2880
         Width           =   2415
      End
      Begin VB.CommandButton btn_generar 
         Caption         =   "Generar"
         Height          =   495
         Left            =   2880
         TabIndex        =   24
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox txt_cantidadNumeros 
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txt_modulo 
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txt_semilla 
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txt_aditiva 
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txt_multiplicador 
         Height          =   375
         Left            =   3600
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.ListBox num_pseudo_list 
         Columns         =   7
         Height          =   2865
         Left            =   120
         TabIndex        =   13
         Top             =   3600
         Width           =   10695
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   9360
         TabIndex        =   2
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label41 
         Caption         =   "Se generará la cantidad necesaria para el problema"
         Height          =   615
         Left            =   7200
         TabIndex        =   85
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese la cantidad de números"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2400
         Width           =   2745
      End
      Begin VB.Label Label4 
         Caption         =   "Ingrese el módulo"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Ingrese la semilla"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese la constante aditiva c"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese la constante multiplicativa a"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Prueba del promedio"
      Height          =   7215
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton btn_promedio_realizar 
         Caption         =   "Realizar"
         Height          =   375
         Left            =   6720
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txt_promedio_alfa 
         Height          =   375
         Left            =   4560
         TabIndex        =   26
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9360
         TabIndex        =   6
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label lbl_promedio_SINO 
         AutoSize        =   -1  'True
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         TabIndex        =   46
         Top             =   4200
         Width           =   285
      End
      Begin VB.Label lbl_promedio_confirmar 
         Caption         =   "Los números están distribuidos uniformemente:"
         Height          =   255
         Left            =   3120
         TabIndex        =   45
         Top             =   4320
         Width           =   4215
      End
      Begin VB.Label lbl_promedio_za2 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   41
         Top             =   3120
         Width           =   120
      End
      Begin VB.Label lbl_promedio_vt 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   40
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label lbl_promedio_alfa 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   39
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label Label17 
         Caption         =   "a/2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   38
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Z"
         Height          =   255
         Left            =   6000
         TabIndex        =   37
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label15 
         Caption         =   "Valor de tablas"
         Height          =   255
         Left            =   6000
         TabIndex        =   36
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Alfa"
         Height          =   255
         Left            =   6000
         TabIndex        =   35
         Top             =   2160
         Width           =   390
      End
      Begin VB.Label lbl_promedio_z0 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   34
         Top             =   3120
         Width           =   120
      End
      Begin VB.Label Label12 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   33
         Top             =   3240
         Width           =   135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Z"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label lbl_promedio_avg 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   31
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Promedio"
         Height          =   255
         Left            =   2520
         TabIndex        =   30
         Top             =   2640
         Width           =   840
      End
      Begin VB.Label lbl_promedio_sum 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sumatoria"
         Height          =   255
         Left            =   2520
         TabIndex        =   28
         Top             =   2160
         Width           =   885
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   9840
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese el grado de significancia"
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   600
         Width           =   2835
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Prueba de la frecuencia"
      Height          =   7215
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Top             =   0
      Width           =   10935
      Begin VB.OptionButton opt_frecuencia_sup 
         Caption         =   "(Lim inf, Lim Sup]"
         Height          =   255
         Left            =   6240
         TabIndex        =   83
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton opt_frecuencia_inferior 
         Caption         =   "[Lim inf, Lim Sup)"
         Height          =   255
         Left            =   6240
         TabIndex        =   82
         Top             =   600
         Width           =   1935
      End
      Begin VB.ListBox list_frecuencia_frecuencias 
         Height          =   2610
         Left            =   5520
         TabIndex        =   59
         Top             =   1920
         Width           =   4335
      End
      Begin VB.ListBox list_frecuencia_numeros 
         Height          =   2610
         Left            =   960
         TabIndex        =   52
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txt_frecuencia_subintervalos 
         Height          =   375
         Left            =   4080
         TabIndex        =   50
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton btn_PruebaFrecuencia 
         Caption         =   "Realizar"
         Height          =   375
         Left            =   8400
         TabIndex        =   49
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_frecuencia_alfa 
         Height          =   375
         Left            =   4080
         TabIndex        =   48
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   9360
         TabIndex        =   9
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Los números están distribuidos uniformemente:"
         Height          =   255
         Left            =   3120
         TabIndex        =   61
         Top             =   6000
         Width           =   4215
      End
      Begin VB.Label lbl_frecuencia_SINO 
         AutoSize        =   -1  'True
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         TabIndex        =   60
         Top             =   5880
         Width           =   285
      End
      Begin VB.Label lbl_frecuencia_xan1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   58
         Top             =   5400
         Width           =   120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "(a,n-1)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4920
         TabIndex        =   57
         Top             =   5520
         Width           =   480
      End
      Begin VB.Label Label20 
         Caption         =   "X²"
         Height          =   255
         Left            =   4800
         TabIndex        =   56
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label lbl_frecuencia_x0 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   55
         Top             =   4920
         Width           =   120
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4920
         TabIndex        =   54
         Top             =   5040
         Width           =   90
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "X²"
         Height          =   255
         Left            =   4800
         TabIndex        =   53
         Top             =   4920
         Width           =   225
      End
      Begin VB.Line Line2 
         X1              =   960
         X2              =   9840
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese el número de intervalos"
         Height          =   255
         Left            =   1080
         TabIndex        =   51
         Top             =   960
         Width           =   2760
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese el grado de significancia"
         Height          =   255
         Left            =   1080
         TabIndex        =   47
         Top             =   480
         Width           =   2835
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Prueba de la distancia"
      Height          =   7215
      Index           =   3
      Left            =   3720
      TabIndex        =   10
      Top             =   0
      Width           =   10935
      Begin VB.ListBox list_distancia_numeros 
         Height          =   3120
         Left            =   960
         TabIndex        =   81
         Top             =   1800
         Width           =   3855
      End
      Begin VB.ListBox list_distancia_tabla 
         Height          =   3120
         Left            =   5040
         TabIndex        =   71
         Top             =   1800
         Width           =   4935
      End
      Begin VB.TextBox txt_distancia_intervalos 
         Height          =   375
         Left            =   3960
         TabIndex        =   69
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_distancia_ls 
         Height          =   375
         Left            =   9000
         TabIndex        =   68
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txt_distancia_li 
         Height          =   375
         Left            =   6480
         TabIndex        =   66
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txt_distancia_alfa 
         Height          =   375
         Left            =   3960
         TabIndex        =   63
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton btn_PruebaDistancia 
         Caption         =   "Realizar"
         Height          =   375
         Left            =   5040
         TabIndex        =   62
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   9360
         TabIndex        =   12
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4920
         TabIndex        =   78
         Top             =   5280
         Width           =   90
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "(a,n-1)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4920
         TabIndex        =   75
         Top             =   5760
         Width           =   480
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "X²"
         Height          =   255
         Left            =   4800
         TabIndex        =   79
         Top             =   5160
         Width           =   225
      End
      Begin VB.Label lbl_distancia_x0 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   77
         Top             =   5160
         Width           =   120
      End
      Begin VB.Label Label30 
         Caption         =   "X²"
         Height          =   255
         Left            =   4800
         TabIndex        =   76
         Top             =   5640
         Width           =   255
      End
      Begin VB.Label lbl_distancia_xan1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   74
         Top             =   5640
         Width           =   120
      End
      Begin VB.Label lbl_distancia_SINO 
         AutoSize        =   -1  'True
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         TabIndex        =   73
         Top             =   6120
         Width           =   285
      End
      Begin VB.Label Label26 
         Caption         =   "Los números están distribuidos uniformemente:"
         Height          =   255
         Left            =   3120
         TabIndex        =   72
         Top             =   6240
         Width           =   4215
      End
      Begin VB.Line Line3 
         X1              =   960
         X2              =   9960
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese el número de distancias"
         Height          =   255
         Left            =   960
         TabIndex        =   70
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Limite superior"
         Height          =   255
         Left            =   7560
         TabIndex        =   67
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Limite inferior"
         Height          =   255
         Left            =   5040
         TabIndex        =   65
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese el grado de significancia"
         Height          =   255
         Left            =   960
         TabIndex        =   64
         Top             =   600
         Width           =   2835
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Aplicación - Parte 1"
      Height          =   7215
      Index           =   5
      Left            =   3720
      TabIndex        =   42
      Top             =   0
      Width           =   10935
      Begin VB.VScrollBar VScroll1 
         Height          =   6135
         LargeChange     =   10
         Left            =   10560
         Max             =   5
         SmallChange     =   10
         TabIndex        =   92
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9360
         TabIndex        =   43
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   44
         Top             =   6720
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -120
         ScaleHeight     =   615
         ScaleWidth      =   11055
         TabIndex        =   95
         Top             =   6600
         Width           =   11055
      End
      Begin VB.PictureBox MFContenedor 
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   120
         ScaleHeight     =   9255
         ScaleWidth      =   10455
         TabIndex        =   86
         Top             =   240
         Width           =   10455
         Begin MSFlexGridLib.MSFlexGrid msfContaminantes 
            Height          =   2580
            Index           =   0
            Left            =   0
            TabIndex        =   87
            Top             =   2640
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4551
            _Version        =   393216
            Rows            =   8
            Cols            =   4
            WordWrap        =   -1  'True
            FormatString    =   ""
         End
         Begin MSFlexGridLib.MSFlexGrid msfMuestrasSangre 
            Height          =   2055
            Left            =   0
            TabIndex        =   93
            Top             =   5640
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   3625
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            WordWrap        =   -1  'True
            FormatString    =   ""
         End
         Begin VB.Label Label46 
            Caption         =   "Resultado de análisis de sangre"
            Height          =   255
            Left            =   0
            TabIndex        =   94
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label Label36 
            Caption         =   $"Proyecto.frx":0004
            Height          =   1335
            Left            =   0
            TabIndex        =   91
            Top             =   480
            Width           =   10215
         End
         Begin VB.Label Label38 
            Caption         =   "Formulación del problema"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   90
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label44 
            Caption         =   "Recolección y procesamiento de datos tomados de la realidad"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   89
            Top             =   1920
            Width           =   6015
         End
         Begin VB.Label Label45 
            Caption         =   "La experiencia indica que la concentración de contaminantes sigue la siguiente distribución de probabilidad"
            Height          =   495
            Left            =   0
            TabIndex        =   88
            Top             =   2280
            Width           =   10335
         End
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Aplicación - Parte 2"
      Height          =   7215
      Index           =   6
      Left            =   3720
      TabIndex        =   96
      Top             =   0
      Width           =   10935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   120
         ScaleHeight     =   6375
         ScaleWidth      =   10695
         TabIndex        =   99
         Top             =   240
         Width           =   10695
         Begin VB.Label Label56 
            Caption         =   $"Proyecto.frx":01D0
            Height          =   495
            Left            =   120
            TabIndex        =   109
            Top             =   3960
            Width           =   10575
         End
         Begin VB.Label Label55 
            Caption         =   $"Proyecto.frx":0268
            Height          =   495
            Left            =   120
            TabIndex        =   108
            Top             =   4560
            Width           =   10575
         End
         Begin VB.Label Label54 
            Caption         =   "Para simular los resultados de las muestras de agua, se requieren 1120 números pseudoaleaorios (20 muestras * 4 puntos * 14 dias)."
            Height          =   495
            Left            =   120
            TabIndex        =   107
            Top             =   3360
            Width           =   10575
         End
         Begin VB.Label Label53 
            Caption         =   $"Proyecto.frx":0301
            Height          =   495
            Left            =   120
            TabIndex        =   106
            Top             =   2760
            Width           =   10455
         End
         Begin VB.Label Label52 
            Caption         =   "Estimación de los parámetros y características operacionales a partir de datos reales"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   2400
            Width           =   8295
         End
         Begin VB.Label Label51 
            Caption         =   $"Proyecto.frx":0388
            Height          =   495
            Left            =   120
            TabIndex        =   104
            Top             =   1680
            Width           =   10455
         End
         Begin VB.Label Label50 
            Caption         =   "2. El agua de los mantos freáticos no es apta para el consumo animal."
            Height          =   255
            Left            =   480
            TabIndex        =   103
            Top             =   1320
            Width           =   7815
         End
         Begin VB.Label Label49 
            Caption         =   "1. El agua de los mantos freáticos es apta para seguir siendo empleada por los animales."
            Height          =   255
            Left            =   480
            TabIndex        =   102
            Top             =   1080
            Width           =   7815
         End
         Begin VB.Label Label48 
            Caption         =   $"Proyecto.frx":0411
            Height          =   495
            Left            =   120
            TabIndex        =   101
            Top             =   480
            Width           =   10455
         End
         Begin VB.Label Label47 
            Caption         =   "Formulación de los modelos matemáticos"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   120
            Width           =   4095
         End
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8280
         TabIndex        =   98
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   97
         Top             =   6720
         Width           =   1455
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Aplicación - Parte 3"
      Height          =   7215
      Index           =   7
      Left            =   3720
      TabIndex        =   110
      Top             =   0
      Width           =   10935
      Begin VB.VScrollBar VScroll2 
         Height          =   6375
         LargeChange     =   10
         Left            =   10440
         Max             =   5
         TabIndex        =   124
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   9360
         TabIndex        =   112
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   111
         Top             =   6720
         Width           =   1455
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   10935
         TabIndex        =   123
         Top             =   6600
         Width           =   10935
      End
      Begin VB.PictureBox MuestrasContenedor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   14700
         Left            =   120
         ScaleHeight     =   14700
         ScaleWidth      =   10335
         TabIndex        =   113
         Top             =   240
         Width           =   10335
         Begin MSFlexGridLib.MSFlexGrid msfMuestrasRes 
            Height          =   780
            Left            =   2040
            TabIndex        =   129
            Top             =   12600
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   1376
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid msfMuestraPunto 
            Height          =   1935
            Index           =   0
            Left            =   2040
            TabIndex        =   117
            Top             =   3120
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   6
            Cols            =   8
         End
         Begin MSFlexGridLib.MSFlexGrid msfMuestras 
            Height          =   1695
            Left            =   2040
            TabIndex        =   116
            Top             =   1080
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   2990
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            FixedRows       =   0
         End
         Begin VB.ListBox lst_Muestras 
            Height          =   12555
            Left            =   120
            TabIndex        =   115
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton btnCalcularMuestras 
            BackColor       =   &H0000FF00&
            Caption         =   "Calcular Análisis de Sangre"
            Height          =   495
            Left            =   120
            MaskColor       =   &H0080FF80&
            Style           =   1  'Graphical
            TabIndex        =   114
            Top             =   120
            Width           =   9975
         End
         Begin MSFlexGridLib.MSFlexGrid msfMuestraPunto 
            Height          =   1935
            Index           =   1
            Left            =   2040
            TabIndex        =   125
            Top             =   5520
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   6
            Cols            =   8
         End
         Begin MSFlexGridLib.MSFlexGrid msfMuestraPunto 
            Height          =   1935
            Index           =   2
            Left            =   2040
            TabIndex        =   126
            Top             =   7920
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   6
            Cols            =   8
         End
         Begin MSFlexGridLib.MSFlexGrid msfMuestraPunto 
            Height          =   1935
            Index           =   3
            Left            =   2040
            TabIndex        =   127
            Top             =   10320
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   6
            Cols            =   8
         End
         Begin VB.Label Label65 
            Caption         =   "Lim. Sup"
            Height          =   255
            Left            =   8880
            TabIndex        =   132
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label64 
            Caption         =   "Lim Inf"
            Height          =   255
            Left            =   7680
            TabIndex        =   131
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label63 
            Caption         =   "Probabilidad"
            Height          =   255
            Left            =   6120
            TabIndex        =   130
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label62 
            Caption         =   "Resultados"
            Height          =   255
            Left            =   2040
            TabIndex        =   128
            Top             =   12360
            Width           =   7215
         End
         Begin VB.Label Label61 
            Caption         =   "Punto 4"
            Height          =   255
            Left            =   2040
            TabIndex        =   122
            Top             =   10080
            Width           =   7215
         End
         Begin VB.Label Label60 
            Caption         =   "Punto 3"
            Height          =   255
            Left            =   2040
            TabIndex        =   121
            Top             =   7680
            Width           =   7215
         End
         Begin VB.Label Label59 
            Caption         =   "Punto 2"
            Height          =   255
            Left            =   2040
            TabIndex        =   120
            Top             =   5280
            Width           =   7215
         End
         Begin VB.Label Label58 
            Caption         =   "Punto 1"
            Height          =   255
            Left            =   2040
            TabIndex        =   119
            Top             =   2880
            Width           =   7215
         End
         Begin VB.Label Label57 
            Caption         =   "Tabla de Probabilidades"
            Height          =   255
            Left            =   2040
            TabIndex        =   118
            Top             =   720
            Width           =   7215
         End
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Aplicación - Parte 4"
      Height          =   7215
      Index           =   8
      Left            =   3720
      TabIndex        =   133
      Top             =   0
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   0
         Left            =   1680
         TabIndex        =   140
         Top             =   2880
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   13
         Left            =   1680
         TabIndex        =   153
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   12
         Left            =   1680
         TabIndex        =   152
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   11
         Left            =   1680
         TabIndex        =   151
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   10
         Left            =   1680
         TabIndex        =   150
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   9
         Left            =   1680
         TabIndex        =   149
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   8
         Left            =   1680
         TabIndex        =   148
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   7
         Left            =   1680
         TabIndex        =   147
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   6
         Left            =   1680
         TabIndex        =   146
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   5
         Left            =   1680
         TabIndex        =   145
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   4
         Left            =   1680
         TabIndex        =   144
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   3
         Left            =   1680
         TabIndex        =   143
         Top             =   2880
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   2
         Left            =   1680
         TabIndex        =   142
         Top             =   2880
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin VB.CommandButton btnAnalisisAgua 
         BackColor       =   &H0000FF00&
         Caption         =   "Calcular muestras de agua"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   360
         Width           =   9135
      End
      Begin VB.ListBox lstNumerosUsados 
         Height          =   5670
         Left            =   120
         TabIndex        =   137
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmbDias 
         Height          =   375
         ItemData        =   "Proyecto.frx":04AD
         Left            =   120
         List            =   "Proyecto.frx":04AF
         TabIndex        =   136
         Text            =   "Dias"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   9360
         TabIndex        =   135
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   134
         Top             =   6720
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid msfAguaPunto 
         Height          =   3615
         Index           =   1
         Left            =   1680
         TabIndex        =   141
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   22
         Cols            =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msfContaminantes 
         Height          =   1500
         Index           =   1
         Left            =   1680
         TabIndex        =   139
         Top             =   840
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2646
         _Version        =   393216
         Rows            =   8
         Cols            =   4
         WordWrap        =   -1  'True
         FormatString    =   ""
      End
   End
   Begin VB.Frame Panel 
      Caption         =   "Aplicación - Portada"
      Height          =   7215
      Index           =   4
      Left            =   3720
      TabIndex        =   168
      Top             =   0
      Width           =   10935
      Begin VB.PictureBox Picture3 
         Height          =   5175
         Left            =   960
         Picture         =   "Proyecto.frx":04B1
         ScaleHeight     =   5115
         ScaleWidth      =   9075
         TabIndex        =   171
         Top             =   480
         Width           =   9135
      End
      Begin VB.CommandButton btn_back 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   170
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton btn_next 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   9360
         TabIndex        =   169
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Sistema de análisis de los mantos freáticos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   172
         Top             =   5880
         Width           =   7695
      End
   End
End
Attribute VB_Name = "vbForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MantosFreaticos As Problema
Dim numeros_pseudoaleatorios() As Double
Dim PanelActual As Integer
Dim Valores As ValoresEstadisticos
Dim Generador As GeneradorPseudoaleatorios
Dim tamaño_muestra As Integer

Dim lngOriginalTop_MFContenedor As Long
Dim lngIncrement_MFContenedor As Long
Dim lngOriginalTop_Muestras As Long
Dim lngIncrement_Muestras As Long

Dim PuntosAptos As Integer
Dim MaximoContaminante As String
Dim FrecuenciaMaximoContaminante As Integer


Private Sub bnt_GenerarDefault_Click()
    numeros_pseudoaleatorios = Generador.GenerarDefault()
    tamaño_muestra = UBound(numeros_pseudoaleatorios)
    num_pseudo_list.Clear
    For i = 0 To tamaño_muestra - 1
        num_pseudo_list.AddItem ((i + 1) & " - " & numeros_pseudoaleatorios(i))
    Next
    Palomazo
End Sub

Private Sub btn_CalcularN_Click()
    txt_multiplicador.Text = 2021
    txt_aditiva = 1
    txt_semilla = 8
    txt_modulo = 320
    txt_cantidadNumeros.Text = 1180
End Sub

Private Sub btn_generar_Click()
    Dim a, c, x0, m, N As Long
    a = Val(txt_multiplicador.Text)
    c = Val(txt_aditiva.Text)
    x0 = Val(txt_semilla.Text)
    m = Val(txt_modulo.Text)
    N = Val(txt_cantidadNumeros.Text)
    Generador.IngresarValores a, c, x0, m, N
    numeros_pseudoaleatorios = Generador.GenerarNumeros()
    tamaño_muestra = UBound(numeros_pseudoaleatorios)
    num_pseudo_list.Clear
    For i = 0 To N - 1
        num_pseudo_list.AddItem ((i + 1) & " - " & numeros_pseudoaleatorios(i))
    Next
    Palomazo
End Sub

Private Sub btn_promedio_realizar_Click()
    Dim alfa, suma, promedio, z0, valor_tablas, z_alfa_entre2 As Double
    alfa = Val(txt_promedio_alfa.Text)
    suma = 0
    For i = 0 To tamaño_muestra - 1
        suma = suma + numeros_pseudoaleatorios(i)
    Next
    promedio = Round(suma / tamaño_muestra, 5)
    z0 = Abs(Round(((promedio - 0.5) * (Sqr(tamaño_muestra)) / (Sqr(1 / 12))), 5))
    valor_tablas = 0.5 - (alfa / 2)
    z_alfa_entre2 = Valores.Normal(alfa / 2)
    
    lbl_promedio_sum.Caption = suma
    lbl_promedio_avg.Caption = promedio
    lbl_promedio_z0.Caption = z0
    lbl_promedio_alfa.Caption = alfa
    lbl_promedio_vt.Caption = valor_tablas
    lbl_promedio_za2.Caption = z_alfa_entre2
    
    If z0 <= z_alfa_entre2 Then
        lbl_promedio_SINO.Caption = "SI"
        Palomazo
    Else
        lbl_promedio_SINO.Caption = "NO"
    End If
End Sub

Private Sub btn_PruebaDistancia_Click()
    list_distancia_tabla.Clear
    list_distancia_numeros.Clear
    
    Dim alfa As Single, Limite_inferior As Single, limite_superior As Single, intervalos As Integer, FE() As Double, FO() As Integer, Chi_Cuadrada_alfa As Double, Chi_Cuadrada_Calculada As Double, Chi_Cuadrada_Parcial() As Double, Dist() As Integer, P_Dist() As Double, Distancia As Integer, Theta As Single, aceptacion As String
        
    alfa = Val(txt_distancia_alfa.Text)
    Limite_inferior = Val(txt_distancia_li.Text)
    limite_superior = Val(txt_distancia_ls.Text)
    Theta = limite_superior - Limite_inferior
    intervalos = Val(txt_distancia_intervalos.Text)
    Distancia = 0
    
    Chi_Cuadrada_alfa = Round(Valores.Chi_Cuadrada(alfa, intervalos - 1), 5)
    ReDim FE(intervalos) As Double, FO(intervalos) As Integer, Chi_Cuadrada_Parcial(intervalos) As Double, Dist(intervalos) As Integer, P_Dist(intervalos) As Double
        
    For i = 0 To intervalos - 1
        Dist(i) = i
        FO(i) = 0
    Next i
    
    list_distancia_numeros.AddItem ("i" & vbTab & "Ri" & vbTab & "Entra" & vbTab & "Dist")
    list_distancia_numeros.AddItem ("-----------------------------------------------------------------------------------------------------------------------------")
    For i = 0 To tamaño_muestra - 1
    aceptacion = "No"
        If numeros_pseudoaleatorios(i) >= Limite_inferior And numeros_pseudoaleatorios(i) <= limite_superior Then
            aceptacion = "SI" & vbTab & Distancia
            For j = 0 To intervalos - 1
                If Distancia = j Or (j = intervalos - 1 And Distancia > intervalos - 1) Then
                    FO(j) = FO(j) + 1
                    Distancia = 0
                End If
            Next j
        Else
            Distancia = Distancia + 1
        End If
        
        list_distancia_numeros.AddItem (i + 1 & vbTab & numeros_pseudoaleatorios(i) & vbTab & aceptacion)
    Next i
    
    For i = 0 To intervalos - 1
        If Distancia = i Or (i = intervalos - 1 And Distancia > i) Then
            FO(i) = FO(i) + 1
            Exit For
        End If
    Next i
    
    list_distancia_tabla.AddItem ("Dist" & vbTab & "FO" & vbTab & "P(Dist)" & vbTab & "FE" & vbTab & "X²")
    list_distancia_tabla.AddItem ("-----------------------------------------------------------------------------------------------------------------------------")
    
    For i = 0 To intervalos - 1
        If i < intervalos - 1 Then
            P_Dist(i) = Theta * ((1 - Theta) ^ i)
        Else
            P_Dist(i) = (1 - Theta) ^ i
        End If
        FE(i) = SumatoriaINT(FO) * P_Dist(i)
        Chi_Cuadrada_Parcial(i) = ((FO(i) - FE(i)) ^ 2) / FE(i)
        list_distancia_tabla.AddItem (i & vbTab & FO(i) & vbTab & Round(P_Dist(i), 5) & vbTab & Round(FE(i), 5) & vbTab & Round(Chi_Cuadrada_Parcial(i), 5))
    Next i
    Chi_Cuadrada_Calculada = Round(SumatoriaDOB(Chi_Cuadrada_Parcial), 5)
    list_distancia_tabla.AddItem ("-----------------------------------------------------------------------------------------------------------------------------")
    list_distancia_tabla.AddItem ("Suma" & vbTab & SumatoriaINT(FO) & vbTab & Round(SumatoriaDOB(P_Dist), 5) & vbTab & Round(SumatoriaDOB(FE), 5) & vbTab & Round(Chi_Cuadrada_Calculada, 5))
    
    lbl_distancia_x0.Caption = Round(Chi_Cuadrada_Calculada, 5)
    lbl_distancia_xan1.Caption = Round(Chi_Cuadrada_alfa, 5)
    
    If Chi_Cuadrada_Calculada <= Chi_Cuadrada_alfa Then
        lbl_distancia_SINO.Caption = "SI"
        Palomazo
    Else
        lbl_distancia_SINO.Caption = "NO"
    End If
End Sub

Private Sub btn_PruebaFrecuencia_Click()
    list_frecuencia_numeros.Clear
    list_frecuencia_frecuencias.Clear
    
    Dim alfa As Double
    Dim numero_subintervalos As Integer
    alfa = Val(txt_frecuencia_alfa.Text)
    numero_subintervalos = Val(txt_frecuencia_subintervalos.Text)
    Dim FE() As Integer, FO() As Integer
    Dim Chi_Cuadrada_Parcial() As Single, Chi_Cuadrada As Double, Chi_Cuadrada_Calculada As Single
    ReDim FE(numero_subintervalos) As Integer, FO(numero_subintervalos) As Integer, Chi_Cuadrada_Parcial(numero_subintervalos) As Single
    
    For i = 0 To numero_subintervalos - 1
        FE(i) = tamaño_muestra / numero_subintervalos
        FO(i) = 0
    Next i
    
    list_frecuencia_numeros.AddItem ("indice" & vbTab & "Ri" & vbTab & "intervalo")
    list_frecuencia_numeros.AddItem "---------------------------------------------------------"
    For i = 0 To tamaño_muestra - 1
        Dim Limite_inferior, limite_superior As Single
        For j = 0 To numero_subintervalos - 1
            Dim subintervalo As String
            Limite_inferior = j * (1 / numero_subintervalos)
            limite_superior = (j + 1) * (1 / numero_subintervalos)
            subintervalo = Limite_inferior & " - " & limite_superior
            
            If opt_frecuencia_inferior.Value Then
                If numeros_pseudoaleatorios(i) >= Limite_inferior And numeros_pseudoaleatorios(i) < limite_superior Then
                    FO(j) = FO(j) + 1
                    list_frecuencia_numeros.AddItem ((i + 1) & vbTab & numeros_pseudoaleatorios(i) & vbTab & subintervalo)
                    Exit For
                End If
            Else
                If numeros_pseudoaleatorios(i) > Limite_inferior And numeros_pseudoaleatorios(i) <= limite_superior Then
                    FO(j) = FO(j) + 1
                    list_frecuencia_numeros.AddItem ((i + 1) & vbTab & numeros_pseudoaleatorios(i) & vbTab & subintervalo)
                    Exit For
                End If
            End If
        Next j
    Next i
    
    list_frecuencia_frecuencias.AddItem ("intervalo" & vbTab & "FE" & vbTab & "FO" & vbTab & "x²")
    list_frecuencia_frecuencias.AddItem "---------------------------------------------------------"
    For i = 0 To numero_subintervalos - 1
        Dim subintervalor As String
        Limite_inferior = i * (1 / numero_subintervalos)
        limite_superior = (i + 1) * (1 / numero_subintervalos)
        subintervalo = Limite_inferior & " - " & limite_superior
        Chi_Cuadrada_Parcial(i) = Round(((FO(i) - FE(i)) ^ 2) / FE(i), 5)
        list_frecuencia_frecuencias.AddItem (subintervalo & vbTab & FE(i) & vbTab & FO(i) & vbTab & Chi_Cuadrada_Parcial(i))
    Next i
    list_frecuencia_frecuencias.AddItem "---------------------------------------------------------"
    list_frecuencia_frecuencias.AddItem ("Suma" & vbTab & SumatoriaINT(FE) & vbTab & SumatoriaINT(FO) & vbTab & SumatoriaSIN(Chi_Cuadrada_Parcial))
    
    Chi_Cuadrada_Calculada = SumatoriaSIN(Chi_Cuadrada_Parcial)
    Chi_Cuadrada = Round(Valores.Chi_Cuadrada(alfa, numero_subintervalos - 1), 5)

    lbl_frecuencia_x0.Caption = Chi_Cuadrada_Calculada
    lbl_frecuencia_xan1.Caption = Chi_Cuadrada
    
    If Chi_Cuadrada_Calculada <= Chi_Cuadrada Then
        lbl_frecuencia_SINO.Caption = "SI"
        Palomazo
    Else
        lbl_frecuencia_SINO.Caption = "NO"
    End If
        
End Sub

Private Sub btn_PruebaSeries_Click()
    flex_series.Clear
    
    Dim alfa As Single, intervalos As Integer, FO() As Integer, FE As Double, tamaño_intervalo As Double, Chi_Cuadrada_Parcial() As Double, Chi_Cuadrada_Calculada As Double, Chi_Cuadrada_alfa As Double, celdas As Integer, pares_ordenados As Integer, Limite_inf As Boolean
    alfa = Val(txt_series_alfa.Text)
    intervalos = Val(txt_series_intervalos.Text)    '4
    celdas = intervalos ^ 2                         '16
    pares_ordenados = tamaño_muestra - 1
    Limite_inf = opt_series_inferior.Value
    
    flex_series.Cols = intervalos
    flex_series.Rows = intervalos
    
    ReDim FO(intervalos, intervalos) As Integer, Chi_Cuadrada_Parcial(intervalos) As Double
    Chi_Cuadrada_alfa = Valores.Chi_Cuadrada(alda, celdas - 1)
    tamaño_intervalo = 1 / intervalos
    FE = pares_ordenados / celdas
    
    For i = 0 To intervalos - 1
        flex_series.RowHeight(i) = flex_series.Height / intervalos
        flex_series.ColWidth(i) = (flex_series.Width - 60) / intervalos
        For j = 0 To intervalos - 1
            FO(i, j) = 0
            flex_series.TextMatrix(i, j) = Str(i + 1 & j + 1)
        Next j
    Next i
    
    primer_valor = numeros_pseudoaleatorios(0)
    flex_series_fo.Cols = 4
    flex_series_fo.Rows = tamaño_muestra + 1
    flex_series_fo.FixedRows = 1
    flex_series_fo.TextMatrix(0, 0) = "Rn"
    flex_series_fo.TextMatrix(0, 1) = "Rn+1"
    flex_series_fo.TextMatrix(0, 2) = "Celda"
    flex_series_fo.TextMatrix(0, 3) = "FO"
    
    For i = 0 To 3
        flex_series_fo.ColWidth(i) = (flex_series_fo.Width - 60) / 4
    Next i
    
    For i = 0 To tamaño_muestra - 1
        rn = numeros_pseudoaleatorios(i)
        If i = tamaño_muestra - 1 Then rnp = primer_valor Else rnp = numeros_pseudoaleatorios(i + 1)
        X = 0
        Y = 0
        
        For j = 0 To intervalos - 1
            x_inf = (1 / intervalos) * j
            x_sup = (1 / intervalos) * (j + 1)
            
            If Limite_inf Then
                If rn >= x_inf And rn < x_sup Then
                    X = j
                End If
            Else
                If rn > x_inf And rn <= x_sup Then
                    X = j
                End If
            End If
            
            For k = 0 To intervalos - 1
                y_inf = (1 / intervalos) * k
                y_sup = (1 / intervalos) * (k + 1)
                
                If Limite_inf Then
                    If rnp >= y_inf And rnp < y_sup Then
                        Y = k
                    End If
                Else
                    If rnp > y_inf And rnp <= y_sup Then
                        Y = k
                    End If
                End If
            Next k
        Next j
        
        FO(X, Y) = FO(X, Y) + 1
        cel = ((X + 1) * 10) + (Y + 1)
        
        flex_series_fo.TextMatrix(i + 1, 0) = rn
        flex_series_fo.TextMatrix(i + 1, 1) = rnp
        flex_series_fo.TextMatrix(i + 1, 2) = Str(cel)
        flex_series_fo.TextMatrix(i + 1, 3) = Str(FO(X, Y))
    Next i
    
    chi_contador = 0
    For i = 0 To intervalos - 1
        For j = 0 To intervalos - 1
            frecuencia = FO(i, j)
            Chi_Cuadrada_Parcial(chi_contador) = ((FO(i, j) - FE) ^ 2) / FE
            flex_series.TextMatrix(i, j) = flex_series.TextMatrix(i, j) & vbCrLf & "FO = " & Str(frecuencia)
            flex_series.TextMatrix(i, j) = flex_series.TextMatrix(i, j) & vbCrLf & "X² = " & Math.Round(Chi_Cuadrada_Parcial(chi_contador), 5)
        Next j
    Next i
    
    lbl_series_FE.Caption = FE
    lbl_series_NP.Caption = pares_ordenados
    lbl_series_TI.Caption = tamaño_intervalo
End Sub

Private Sub SetUsedNumbers(dia As Integer)
    final = (dia * 80) + 60
    inicial = final - 80
    
    lstNumerosUsados.Clear
    For i = inicial To final - 1
        lstNumerosUsados.AddItem (i + 1 & " - " & Round(numeros_pseudoaleatorios(i), 2))
    Next i
    
End Sub

Private Sub btnAnalisisAgua_Click()
    Ri = 0
    indice = 60
    
    ContadorSC = 0
    ContadorEM = 0
    ContadorRP = 0
    ContadorSF = 0
    ContadorAC = 0
    ContadorFF = 0
    ContadorOX = 0
    ValMax = 0
    Max = ""
    For i = 0 To 13     'Dias'
        For j = 0 To 19     'Muestras'
            Resultado = ""
            For k = 0 To 3      'Puntos'
                Ri = numeros_pseudoaleatorios(indice)
                msfAguaPunto(i).TextMatrix(j + 1, ((k + 1) * 2) - 1) = Ri
                
                For l = 0 To 6
                    Limite_inferior = Val(msfContaminantes(1).TextMatrix(l + 1, 2))
                    limite_superior = Val(msfContaminantes(1).TextMatrix(l + 1, 3))
                    If Ri >= Limite_inferior And Ri < limite_superior Then
                        Resultado = msfContaminantes(1).TextMatrix(l + 1, 0)
                        Resultado = Mid(Resultado, Len(Resultado) - 2, 2)
                        Exit For
                    End If
                Next l
                
                Select Case Resultado
                    Case "SC"
                        ContadorSC = ContadorSC + 1
                        If ContadorSC > ValMax Then
                            ValMax = ContadorSC
                            Max = "Substancias Coloidales"
                        End If
                    Case "EM"
                        ContadorEM = ContadorEM + 1
                        If ContadorEM > ValMax Then
                            ValMax = ContadorEM
                            Max = "Exceso de mercurio"
                        End If
                    Case "RP"
                        ContadorRP = ContadorRP + 1
                        If ContadorRP > ValMax Then
                            ValMax = ContadorRP
                            Max = "Resíduos petroquímicos"
                        End If
                    Case "SF"
                        ContadorSF = ContadorSF + 1
                        If ContadorSF > ValMax Then
                            ValMax = ContadorSF
                            Max = "Sulfatos"
                        End If
                    Case "AC"
                        ContadorAC = ContadorAC + 1
                        If ContadorAC > ValMax Then
                            ValMax = ContadorAC
                            Max = "Ácido clorhídrico"
                        End If
                    Case "FF"
                        ContadorFF = ContadorFF + 1
                        If ContadorFF > ValMax Then
                            ValMax = ContadorFF
                            Max = "Fosfátos"
                        End If
                    Case "OX"
                        ContadorOX = ContadorOX + 1
                        If ContadorOX > ValMax Then
                            ValMax = ContadorOX
                            Max = "Óxidos"
                        End If
                End Select
                msfAguaPunto(i).TextMatrix(j + 1, (k + 1) * 2) = Resultado
                indice = indice + 1
            Next k
        Next j
    Next i
    
    MaximoContaminante = Max
    FrecuenciaMaximoContaminante = ValMax
    
    lblPuntosAb.Caption = PuntosAptos
    If PuntosAptos >= 3 Then
        lblPredomina.Caption = "SI"
        lblApta.Caption = "SI"
    Else
        lblPredomina.Caption = "NO"
        lblApta.Caption = "NO"
    End If
    
    lblContaminante.Caption = MaximoContaminante
    lblFrecuencia.Caption = FrecuenciaMaximoContaminante
    
End Sub

Private Sub btnCalcularMuestras_Click()
    For i = 0 To 59
        lst_Muestras.AddItem ((1 + i) & " - " & Round(numeros_pseudoaleatorios(i), 2))
    Next i
      
    'Llenar las casillas con numeros pseudoaleatorios'
    
    Ri = 0
    indice = 0
    
    For i = 0 To 3  'Puntos'
        contadorAPTA = 0 'Apta'
        For j = 0 To 4  'Animales'
            contadorRN = 0  'Estado de rango normal'
            For k = 0 To 2 'Etapas'
                Resultado = ""
                Ri = numeros_pseudoaleatorios(indice)
                msfMuestraPunto(i).TextMatrix(j + 1, ((k + 1) * 2) - 1) = Ri
                For l = 0 To 4
                    Limite_inferior = Val(msfMuestras.TextMatrix(l, 2))
                    limite_superior = Val(msfMuestras.TextMatrix(l, 3))
                    If Ri >= Limite_inferior And Ri < limite_superior Then
                        Resultado = msfMuestras.TextMatrix(l, 0)
                        Resultado = Mid(Resultado, Len(Resultado) - 2, 2)
                        Exit For
                    End If
                Next l
                
                If Resultado = "RN" Then
                    contadorRN = contadorRN + 1
                End If
                
                msfMuestraPunto(i).TextMatrix(j + 1, (k + 1) * 2) = Resultado
                indice = indice + 1
            Next k
            If contadorRN >= 2 Then
                msfMuestraPunto(i).TextMatrix(j + 1, 7) = "Apta"
                contadorAPTA = contadorAPTA + 1
            Else
                msfMuestraPunto(i).TextMatrix(j + 1, 7) = "No apta"
            End If
        Next j
        If contadorAPTA > 2 Then
            msfMuestrasRes.TextMatrix(1, i) = "Apta"
            PuntosAptos = PuntosAptos + 1
        Else
            msfMuestrasRes.TextMatrix(1, i) = "No Apta"
        End If
    Next i
End Sub

Private Sub cmbDias_Click()
    SetUsedNumbers (cmbDias.ListIndex + 1)
    For i = 0 To 13
        msfAguaPunto(i).Visible = False
    Next i
    msfAguaPunto(cmbDias.ListIndex).Visible = True
    msfAguaPunto(cmbDias.ListIndex).ZOrder (0)
End Sub

Private Sub Form_Activate()
    txt_multiplicador.SetFocus
End Sub

Private Sub Form_Load()
    Ajustar_TODOList
    Siguiente_Panel (0)
    Set Valores = New ValoresEstadisticos
    Set Generador = New GeneradorPseudoaleatorios
    Set MantosFreaticos = New Problema
    SetProblemValues
    PuntosAptos = 0
    lngOriginalTop_MFContenedor = MFContenedor.Top
    lngIncrement_MFContenedor = (MFContenedor.Height - Panel(5).Height) / VScroll1.Max
    
    lngOriginalTop_Muestras = MuestrasContenedor.Top
    lngIncrement_Muestras = (MuestrasContenedor.Height - Panel(7).Height) / VScroll2.Max
End Sub

Private Sub Ajustar_TODOList()
    todo_list.Top = 0
    todo_list.Left = 0
    todo_list.Height = ScaleHeight
    todo_list.Width = 3600
    
    For i = 0 To Panel.Count - 1
        
    
        todo_list.AddItem Panel(i), i
    Next
End Sub

Private Sub Siguiente_Panel(i As Integer)
    PanelActual = i
    For j = 0 To Panel.Count - 1
        Panel(j).Visible = False
    Next
    Panel(i).Visible = True
    Panel(i).ZOrder (0)
    Panel(i).Left = todo_list.Width + 100
    Panel(i).Top = 0
    Panel(i).Width = ScaleWidth - todo_list.Width - 200
    Panel(i).Height = ScaleHeight - 100
    btn_next(i).Top = Panel(i).Height - btn_next(i).Height - 100
    btn_next(i).Left = Panel(i).Width - btn_next(i).Width - 100
    btn_back(i).Top = Panel(i).Height - btn_next(i).Height - 100
    btn_back(i).Left = 100
End Sub

Private Sub btn_next_Click(Index As Integer)
    If (Panel.Count > Index + 1) Then
        Siguiente_Panel (Index + 1)
        Palomazo
    Else
        MsgBox "Proyecto terminado"
    End If
End Sub

Private Sub btn_back_Click(Index As Integer)
    If (Index > 0) Then
        Siguiente_Panel (Index - 1)
    End If
End Sub

Private Sub todo_list_Click()
    Siguiente_Panel (todo_list.ListIndex)
End Sub

Private Sub Palomazo()
    todo_list.Selected(PanelActual) = True
End Sub

Private Function SumatoriaINT(arreglo() As Integer) As Integer
    Dim temp As Integer
    temp = 0
    For i = 0 To UBound(arreglo) - 1
        temp = temp + arreglo(i)
    Next i
    SumatoriaINT = temp
End Function

Private Function SumatoriaSIN(arreglo() As Single) As Single
    Dim temp As Single
    temp = 0
    For i = 0 To UBound(arreglo) - 1
        temp = temp + arreglo(i)
    Next i
    SumatoriaSIN = temp
End Function

Private Function SumatoriaDOB(arreglo() As Double) As Double
    Dim temp As Double
    temp = 0
    For i = 0 To UBound(arreglo) - 1
        temp = temp + arreglo(i)
    Next i
    SumatoriaDOB = temp
End Function

Private Sub SetProblemValues()
    'Tablas de contaminantes y muestras de sangre'
    msfContaminantes(0).ColWidth(0) = (msfContaminantes(0).Width / 2) - 120
    msfContaminantes(1).ColWidth(0) = (msfContaminantes(1).Width / 2) - 400
    msfMuestrasSangre.ColWidth(0) = (msfMuestrasSangre.Width / 2) - 120
    msfMuestras.ColWidth(0) = (msfMuestras.Width / 2) - 120
    
    For i = 1 To 3
        msfContaminantes(0).ColWidth(i) = msfContaminantes(0).Width / 6
        msfContaminantes(1).ColWidth(i) = msfContaminantes(1).Width / 6
        msfMuestrasSangre.ColWidth(i) = msfMuestrasSangre.Width / 6
        msfMuestras.ColWidth(i) = msfMuestras.Width / 6
    Next i
    
    For i = 0 To 1
        msfContaminantes(i).TextMatrix(0, 0) = "Concentración de contaminantes"
        msfContaminantes(i).TextMatrix(0, 1) = "Prob"
        msfContaminantes(i).TextMatrix(0, 2) = "Lim. Inf"
        msfContaminantes(i).TextMatrix(0, 3) = "Lim. Sup"
    Next i
    
    acum = 0
    For i = 0 To 6
        Descripcion = MantosFreaticos.GetContaminantes(i)
        Probabilidad = MantosFreaticos.GetProbabilidad_Contaminacion(i)
        Liminf = acum
        acum = acum + Probabilidad
        Limsup = acum
        
        For j = 0 To 1
            msfContaminantes(j).TextMatrix(i + 1, 0) = Descripcion
            msfContaminantes(j).TextMatrix(i + 1, 1) = Probabilidad
            msfContaminantes(j).TextMatrix(i + 1, 2) = Liminf
            msfContaminantes(j).TextMatrix(i + 1, 3) = Limsup
        Next j
    Next i
    
    msfMuestrasSangre.TextMatrix(0, 0) = "Resultados de analisis de sangre"
    msfMuestrasSangre.TextMatrix(0, 1) = "Prob"
    msfMuestrasSangre.TextMatrix(0, 2) = "Lim. Inf"
    msfMuestrasSangre.TextMatrix(0, 3) = "Lim. Sup"
    
    acum = 0
    For i = 0 To 4
        Descripcion = MantosFreaticos.GetAnalisis_Sangre(i)
        Probabilidad = MantosFreaticos.GetProbabilidad_Analisis_Sangre(i)
        Liminf = acum
        acum = acum + Probabilidad
        Limsup = acum
        
        msfMuestrasSangre.TextMatrix(i + 1, 0) = Descripcion
        msfMuestrasSangre.TextMatrix(i + 1, 1) = Probabilidad
        msfMuestrasSangre.TextMatrix(i + 1, 2) = Liminf
        msfMuestrasSangre.TextMatrix(i + 1, 3) = Limsup
        
        msfMuestras.TextMatrix(i, 0) = Descripcion
        msfMuestras.TextMatrix(i, 1) = Probabilidad
        msfMuestras.TextMatrix(i, 2) = Liminf
        msfMuestras.TextMatrix(i, 3) = Limsup
    Next i
    
    'Tablas de los puntos del abrevadero'
    For i = 0 To 3
        For j = 0 To 7
            msfMuestraPunto(i).ColWidth(j) = (msfMuestraPunto(i).Width - 80) / 8
        Next j
        msfMuestraPunto(i).TextMatrix(0, 0) = "Animal"
        msfMuestraPunto(i).TextMatrix(0, 1) = "Ri"
        msfMuestraPunto(i).TextMatrix(0, 2) = "Res"
        msfMuestraPunto(i).TextMatrix(0, 3) = "Ri"
        msfMuestraPunto(i).TextMatrix(0, 4) = "Res"
        msfMuestraPunto(i).TextMatrix(0, 5) = "Ri"
        msfMuestraPunto(i).TextMatrix(0, 6) = "Res"
        msfMuestraPunto(i).TextMatrix(0, 7) = "Con"
        
        For j = 1 To 5
            msfMuestraPunto(i).TextMatrix(j, 0) = j
        Next j
    Next i
    
    'Tabla de resultados'
    For i = 0 To 3
        msfMuestrasRes.ColWidth(i) = (msfMuestrasRes.Width - 120) / 4
        msfMuestrasRes.TextMatrix(0, i) = "Punto " & i + 1
    Next i
    
    'Tabla del analisis de agua
    For i = 0 To 13
    
        msfAguaPunto(i).Width = 9135
        
        For j = 0 To 8
            msfAguaPunto(i).ColWidth(j) = (msfAguaPunto(i).Width / 9) - 40
        Next j
        
        msfAguaPunto(i).TextMatrix(0, 0) = "Muestra"
        msfAguaPunto(i).TextMatrix(0, 1) = "Ri"
        msfAguaPunto(i).TextMatrix(0, 2) = "Res"
        msfAguaPunto(i).TextMatrix(0, 3) = "Ri"
        msfAguaPunto(i).TextMatrix(0, 4) = "Res"
        msfAguaPunto(i).TextMatrix(0, 5) = "Ri"
        msfAguaPunto(i).TextMatrix(0, 6) = "Res"
        msfAguaPunto(i).TextMatrix(0, 7) = "Ri"
        msfAguaPunto(i).TextMatrix(0, 8) = "Res"
        
        For j = 0 To 19
            msfAguaPunto(i).TextMatrix(j + 1, 0) = j + 1
        Next j
    Next i
    
    'Combo de dias
    For i = 0 To 13
        cmbDias.AddItem ("Dia " & i + 1)
    Next i
    
End Sub

Private Sub VScroll1_Change()
    MFContenedor.Top = lngOriginalTop_MFContenedor - (VScroll1.Value * lngIncrement_MFContenedor)
End Sub

Private Sub VScroll2_Change()
    MuestrasContenedor.Top = lngOriginalTop_Muestras - (VScroll2.Value * lngIncrement_Muestras)
End Sub

