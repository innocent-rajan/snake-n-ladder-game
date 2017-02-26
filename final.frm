VERSION 5.00
Begin VB.Form Snake 
   Caption         =   "Snake and Ladder"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton roll3 
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   125
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton roll4 
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   122
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton roll2 
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   119
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      TabIndex        =   106
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton roll1 
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox p3 
      Caption         =   "     Three players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   2
      Top             =   7680
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select number of players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   11160
      TabIndex        =   0
      Top             =   6120
      Width           =   3255
      Begin VB.CheckBox p4 
         Caption         =   "     Four players"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   110
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CheckBox p2 
         Caption         =   "     Two players"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   23
      Left            =   7200
      TabIndex        =   31
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   28
      Left            =   3600
      TabIndex        =   36
      Top             =   8400
      Width           =   375
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   18
      Left            =   6600
      TabIndex        =   26
      Top             =   9120
      Width           =   375
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   57
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "57"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   54
      Left            =   3600
      TabIndex        =   61
      Top             =   5520
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   60
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   51
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   60
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   51
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   60
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   51
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   60
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   51
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   57
      Left            =   1440
      TabIndex        =   64
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "51"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   48
      Left            =   7920
      TabIndex        =   55
      Top             =   5520
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   59
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   58
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   57
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   56
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   55
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   54
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   53
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   52
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   59
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   58
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   56
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   55
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   54
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   53
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   52
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   59
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   58
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   57
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   56
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   55
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   54
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   53
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   52
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   59
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   58
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   57
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   56
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   55
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   54
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   53
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   52
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "59"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   56
      Left            =   2160
      TabIndex        =   63
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "58"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   55
      Left            =   2880
      TabIndex        =   62
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "56"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   53
      Left            =   4320
      TabIndex        =   60
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "55"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   52
      Left            =   5040
      TabIndex        =   59
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   51
      Left            =   5760
      TabIndex        =   58
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   50
      Left            =   6480
      TabIndex        =   57
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "52"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   49
      Left            =   7200
      TabIndex        =   56
      Top             =   5520
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   61
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   61
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   61
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   61
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "61"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   67
      Left            =   1440
      TabIndex        =   107
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   65
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   64
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   63
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   62
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   65
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   64
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   63
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   62
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   65
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   64
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   63
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   62
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   65
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   64
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   63
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   62
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "62"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   66
      Left            =   2160
      TabIndex        =   73
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "63"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   65
      Left            =   2880
      TabIndex        =   72
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "64"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   64
      Left            =   3600
      TabIndex        =   71
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "65"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   63
      Left            =   4320
      TabIndex        =   70
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   70
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   70
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   70
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   70
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   69
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   68
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   67
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   66
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   69
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   68
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   67
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   66
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   69
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   68
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   67
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   66
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   69
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   68
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   67
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   66
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "66"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   62
      Left            =   5040
      TabIndex        =   69
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "67"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   61
      Left            =   5760
      TabIndex        =   68
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "68"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   60
      Left            =   6480
      TabIndex        =   67
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "69"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   59
      Left            =   7200
      TabIndex        =   66
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   58
      Left            =   8040
      TabIndex        =   65
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   80
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   71
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   80
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   71
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   80
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   71
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   80
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   71
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "80"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   77
      Left            =   1440
      TabIndex        =   83
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "71"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   68
      Left            =   7920
      TabIndex        =   74
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   79
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   78
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   77
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   76
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   75
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   74
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   73
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   72
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   79
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   78
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   77
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   76
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   75
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   74
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   73
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   72
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   79
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   79
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   78
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   77
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   76
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   75
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   74
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   73
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   72
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   78
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   77
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   76
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   75
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   74
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   73
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   72
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "79"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   76
      Left            =   2160
      TabIndex        =   82
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "78"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   75
      Left            =   2880
      TabIndex        =   81
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "77"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   74
      Left            =   3600
      TabIndex        =   80
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "76"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   73
      Left            =   4320
      TabIndex        =   79
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   72
      Left            =   5040
      TabIndex        =   78
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "74"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   71
      Left            =   5760
      TabIndex        =   77
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "73"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   70
      Left            =   6480
      TabIndex        =   76
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "72"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   69
      Left            =   7200
      TabIndex        =   75
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   90
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   81
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   90
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   81
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   90
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   81
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   90
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   81
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "81"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   87
      Left            =   1440
      TabIndex        =   93
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "90"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   78
      Left            =   7920
      TabIndex        =   84
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   89
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   88
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   87
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   86
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   85
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   84
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   83
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   82
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   89
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   88
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   87
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   86
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   85
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   84
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   83
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   82
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   89
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   88
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   87
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   86
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   85
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   84
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   83
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   82
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   89
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   88
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   87
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   86
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   85
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   84
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   83
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   82
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "82"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   86
      Left            =   2160
      TabIndex        =   92
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "83"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   85
      Left            =   2880
      TabIndex        =   91
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "84"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   84
      Left            =   3600
      TabIndex        =   90
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "85"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   83
      Left            =   4320
      TabIndex        =   89
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "86"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   82
      Left            =   5040
      TabIndex        =   88
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "87"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   81
      Left            =   5760
      TabIndex        =   87
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   80
      Left            =   6480
      TabIndex        =   86
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "89"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   79
      Left            =   7200
      TabIndex        =   85
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "93"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   90
      Left            =   6480
      TabIndex        =   96
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   93
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   97
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   95
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   94
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   93
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   92
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   91
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   97
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   96
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   95
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   94
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   93
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   92
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   91
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   97
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   96
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   95
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   94
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   93
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   92
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   91
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   97
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   96
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   95
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   94
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   92
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   91
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "97"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   94
      Left            =   3600
      TabIndex        =   100
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "96"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   93
      Left            =   4320
      TabIndex        =   99
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "95"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   92
      Left            =   5040
      TabIndex        =   98
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "94"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   91
      Left            =   5760
      TabIndex        =   97
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "92"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   89
      Left            =   7200
      TabIndex        =   95
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "91"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   88
      Left            =   7920
      TabIndex        =   94
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   96
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   100
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   99
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   98
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   100
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   99
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   98
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   100
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   99
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   98
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   100
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   99
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   98
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   97
      Left            =   1320
      TabIndex        =   103
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   96
      Left            =   2160
      TabIndex        =   102
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "98"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   95
      Left            =   2880
      TabIndex        =   101
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   50
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   41
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   40
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   31
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   50
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   41
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   40
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   31
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   50
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   41
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   40
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   31
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   50
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   41
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   40
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   31
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "41"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   47
      Left            =   1440
      TabIndex        =   54
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   38
      Left            =   7920
      TabIndex        =   45
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   37
      Left            =   1440
      TabIndex        =   44
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   11
      Left            =   7920
      TabIndex        =   20
      Top             =   6960
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   49
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   48
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   47
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   46
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   45
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   44
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   43
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   42
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   39
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   38
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   37
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   36
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   35
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   34
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   33
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   32
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   49
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   48
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   47
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   46
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   45
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   44
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   43
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   42
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   39
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   38
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   37
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   36
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   35
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   34
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   33
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   32
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   49
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   48
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   47
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   46
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   45
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   44
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   43
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   42
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   39
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   38
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   37
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   36
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   35
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   34
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   33
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   32
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   49
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   48
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   47
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   46
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   45
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   44
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   43
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   42
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   39
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   38
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   37
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   36
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   35
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   34
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   33
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   32
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "42"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   46
      Left            =   2160
      TabIndex        =   53
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "43"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   45
      Left            =   2880
      TabIndex        =   52
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "44"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   44
      Left            =   3600
      TabIndex        =   51
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   43
      Left            =   4320
      TabIndex        =   50
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "46"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   42
      Left            =   5040
      TabIndex        =   49
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   41
      Left            =   5760
      TabIndex        =   48
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "48"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   40
      Left            =   6480
      TabIndex        =   47
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "49"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   39
      Left            =   7200
      TabIndex        =   46
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "39"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   34
      Left            =   2160
      TabIndex        =   41
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   30
      Left            =   2880
      TabIndex        =   38
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "33"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   24
      Left            =   6480
      TabIndex        =   32
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   21
      Left            =   4320
      TabIndex        =   29
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "37"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   15
      Left            =   3600
      TabIndex        =   24
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   14
      Left            =   5040
      TabIndex        =   23
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   13
      Left            =   5760
      TabIndex        =   22
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   12
      Left            =   7200
      TabIndex        =   21
      Top             =   6960
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   30
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   30
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   30
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   30
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   36
      Left            =   1440
      TabIndex        =   43
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   29
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   28
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   27
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   26
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   25
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   24
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   28
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   27
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   26
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   25
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   24
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   29
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   28
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   27
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   26
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   25
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   24
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   29
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   28
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   27
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   26
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   25
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   24
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   33
      Left            =   2160
      TabIndex        =   108
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   29
      Left            =   2880
      TabIndex        =   37
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   25
      Left            =   5760
      TabIndex        =   33
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   20
      Left            =   3600
      TabIndex        =   28
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   19
      Left            =   5040
      TabIndex        =   27
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   10
      Left            =   4320
      TabIndex        =   19
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   9
      Left            =   6480
      TabIndex        =   18
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   8
      Left            =   7920
      TabIndex        =   17
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   16
      Left            =   7200
      TabIndex        =   109
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   35
      Left            =   1440
      TabIndex        =   42
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   22
      Left            =   7920
      TabIndex        =   30
      Top             =   8400
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   19
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   19
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   19
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   19
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   99
      Left            =   4320
      TabIndex        =   105
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   98
      Left            =   6480
      TabIndex        =   104
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   32
      Left            =   2160
      TabIndex        =   40
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   26
      Left            =   5040
      TabIndex        =   34
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   17
      Left            =   5760
      TabIndex        =   25
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   10
      Top             =   8400
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   27
      Left            =   4440
      TabIndex        =   35
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   6
      Left            =   3720
      TabIndex        =   15
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   5
      Left            =   5160
      TabIndex        =   14
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   5880
      TabIndex        =   13
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   7320
      TabIndex        =   12
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   8040
      TabIndex        =   11
      Top             =   9120
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   31
      Left            =   2280
      TabIndex        =   39
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   7
      Left            =   3000
      TabIndex        =   16
      Top             =   9120
      Width           =   375
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   9120
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   7230
      Left            =   1320
      Picture         =   "final.frx":0000
      Top             =   2520
      Width           =   7260
   End
   Begin VB.Label n3 
      Caption         =   "The Number is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   127
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape c31 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   13800
      Shape           =   3  'Circle
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c35 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c37 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   14160
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c32 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   13800
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c36 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c34 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c33 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   13800
      Shape           =   3  'Circle
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label m3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   126
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label n4 
      Caption         =   "The Number is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   124
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape d4 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   13680
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape c21 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c25 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   10320
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c27 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9960
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c22 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c26 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   10320
      Shape           =   3  'Circle
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c24 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   10320
      Shape           =   3  'Circle
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c23 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label m4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14520
      TabIndex        =   123
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label n1 
      Caption         =   "The Number is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   121
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape c11 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   13800
      Shape           =   3  'Circle
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c15 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c17 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   14160
      Shape           =   3  'Circle
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c12 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   13800
      Shape           =   3  'Circle
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c16 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c14 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c13 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   13800
      Shape           =   3  'Circle
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label m1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   120
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   118
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   117
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   116
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   115
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7440
      TabIndex        =   114
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5400
      TabIndex        =   113
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3480
      TabIndex        =   112
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1680
      TabIndex        =   111
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   29
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   7800
      Width           =   255
   End
   Begin VB.Shape s4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   960
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   600
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape s2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   720
      Shape           =   3  'Circle
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape s1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   720
      Shape           =   3  'Circle
      Top             =   9000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   99
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   98
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   97
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   96
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   95
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   94
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   93
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   92
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   91
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   90
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   89
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   88
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   87
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   86
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   85
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   84
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   83
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   82
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   81
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   80
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   79
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   78
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   77
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   76
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   75
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   74
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   73
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   72
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   71
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   70
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   69
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   68
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   67
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   66
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   65
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   64
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   63
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   62
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   61
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   60
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   59
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   58
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   57
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   56
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   55
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   54
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   53
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   52
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   51
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   50
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   49
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   48
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   47
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   46
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   45
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   44
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   43
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   42
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   41
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   40
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   39
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   38
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   37
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   36
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   35
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   34
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   33
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   32
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   31
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   30
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   29
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   28
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   27
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   26
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   25
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   24
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   23
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   22
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   21
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   20
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   19
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   18
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   17
      Left            =   5640
      Shape           =   5  'Rounded Square
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   16
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   15
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   14
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   13
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   12
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   11
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   10
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   9
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   8
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   7
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   6
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   5
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   4
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   4920
      Shape           =   5  'Rounded Square
      Top             =   7560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   2
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   0
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label m2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape c3 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c4 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   10320
      Shape           =   3  'Circle
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c6 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   10320
      Shape           =   3  'Circle
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c2 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c7 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9960
      Shape           =   3  'Circle
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c5 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   10320
      Shape           =   3  'Circle
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape c1 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape d2 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   13680
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label n2 
      Caption         =   "The Number is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape d1 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape d3 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Snake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim no1, no2, no3, no4
Dim t1
Dim a, b, c, d As Integer
Dim q, r As Integer
Dim a1, a2, a3, q4 As Integer
Dim ax As Integer
Private Sub Command1_Click()
q = MsgBox("Are you sure you want to quit", vbYesNo)
If q = vbYes Then
    End
End If
End Sub

Private Sub Command2_Click()
MsgBox "Snake And Ladder Made by Rajan Girsa"
End Sub

Private Sub Command3_Click()
MsgBox "E-mail Id : jaatboyrajan@gmail.com"
End Sub

Private Sub Command4_Click()
r = MsgBox("Are you sure you want to restart ", vbYesNo)
If r = vbYes Then
Unload Snake
Snake.Show
End If
End Sub

Private Sub Form_Load()
a = 0
b = 0
c = 0
d = 0
x = 0
y = 0
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
c4.Visible = False
c5.Visible = False
c6.Visible = False
c7.Visible = False
d1.Visible = False
roll1.Visible = False
Command4.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
roll1.Visible = False
roll2.Visible = False
roll3.Visible = False
roll4.Visible = False
End Sub

Private Sub p2_Click()
n1.Visible = True
n2.Visible = True
m1.Visible = True
m2.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
p2.Visible = False
p3.Visible = False
Frame1.Visible = False
Label6.Visible = False
c1.Visible = True
c2.Visible = True
c3.Visible = True
c4.Visible = True
c5.Visible = True
c6.Visible = True
c7.Visible = True
d1.Visible = True
roll1.Visible = True
c11.Visible = True
c12.Visible = True
c13.Visible = True
c14.Visible = True
c15.Visible = True
c16.Visible = True
c17.Visible = True
d2.Visible = True
Command4.Visible = True
s1(a).Visible = True
s2(b).Visible = True
Label3.Visible = True
Label9.Visible = True
Label4.Visible = True
Label10.Visible = True
a1 = InputBox("Enter player # 1's name. ")
Label3.Caption = a1
a2 = InputBox("Enter player # 2's name. ")
Label4.Caption = a2
If Label3.Caption = "" Then
a1 = InputBox("Please Enter the name of player # 1. ")
Label3.Caption = a1
If Label4.Caption = "" Then
a2 = InputBox("Please Enter the name of player # 2. ")
Label4.Caption = a2
End If
End If
Label9.Visible = True
Label10.Visible = True
Label9.Caption = Str(a)
Label10.Caption = Str(b)
End Sub

Private Sub p3_Click()
n1.Visible = True
n2.Visible = True
n3.Visible = True
m1.Visible = True
m2.Visible = True
m3.Visible = True
c1.Visible = True
c2.Visible = True
c3.Visible = True
c4.Visible = True
c5.Visible = True
c6.Visible = True
c7.Visible = True
d1.Visible = True
roll1.Visible = True
c11.Visible = True
c12.Visible = True
c13.Visible = True
c14.Visible = True
c15.Visible = True
c16.Visible = True
c17.Visible = True
d2.Visible = True
c21.Visible = True
c22.Visible = True
c23.Visible = True
c24.Visible = True
c25.Visible = True
c26.Visible = True
c27.Visible = True
d3.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
p2.Visible = False
p3.Visible = False
Frame1.Visible = False
Label6.Visible = False
Command4.Visible = True
s1(a).Visible = True
s2(b).Visible = True
s3(c).Visible = True
Label3.Visible = True
Label9.Visible = True
Label4.Visible = True
Label10.Visible = True
Label5.Visible = True
Label11.Visible = True
a1 = InputBox("Enter the name of player # 1. ")
Label3.Caption = a1
a2 = InputBox("Enter the name of player # 2. ")
Label4.Caption = a2
a3 = InputBox("Enter the name of player # 3. ")
Label5.Caption = a3
If Label3.Caption = "" Then
    a1 = InputBox("Please Enter the name of player # 1. ")
Label3.Caption = a1
If Label4.Caption = "" Then
    a1 = InputBox("Please Enter the name of player # 2. ")
Label4.Caption = a2
If Label5.Caption = "" Then
    a1 = InputBox("Please Enter the name of player # 3. ")
Label5.Caption = a3
End If
End If
End If
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label9.Caption = Str(a)
Label10.Caption = Str(b)
Label11.Caption = Str(c)
End Sub

Private Sub p4_Click()
n1.Visible = True
n2.Visible = True
n3.Visible = True
n4.Visible = True
m1.Visible = True
m2.Visible = True
m3.Visible = True
m4.Visible = True
c1.Visible = True
c2.Visible = True
c3.Visible = True
c4.Visible = True
c5.Visible = True
c6.Visible = True
c7.Visible = True
d1.Visible = True
roll1.Visible = True
c11.Visible = True
c12.Visible = True
c13.Visible = True
c14.Visible = True
c15.Visible = True
c16.Visible = True
c17.Visible = True
d2.Visible = True
c21.Visible = True
c22.Visible = True
c23.Visible = True
c24.Visible = True
c25.Visible = True
c26.Visible = True
c27.Visible = True
d3.Visible = True
c31.Visible = True
c32.Visible = True
c33.Visible = True
c34.Visible = True
c35.Visible = True
c36.Visible = True
c37.Visible = True
d4.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
p2.Visible = False
p3.Visible = False
Frame1.Visible = False
Label6.Visible = False
c1.Visible = True
c2.Visible = True
c3.Visible = True
c4.Visible = True
c5.Visible = True
c6.Visible = True
c7.Visible = True
d1.Visible = True
roll1.Visible = True
Command4.Visible = True
s1(a).Visible = True
s2(b).Visible = True
s3(c).Visible = True
s4(d).Visible = True
Label3.Visible = True
Label9.Visible = True
Label4.Visible = True
Label10.Visible = True
Label5.Visible = True
Label11.Visible = True
Label6.Visible = True
Label12.Visible = True
a1 = InputBox("Enter player # 1's name.")
Label3.Caption = a1
a2 = InputBox("Enter player # 2's name.")
Label4.Caption = a2
a3 = InputBox("Enter player # 3's name.")
Label5.Caption = a3
a4 = InputBox("Enter player # 4's name.")
Label6.Caption = a4
If Label3.Caption = "" Then
    a1 = InputBox("Please Enter the name of player # 1. ")
Label3.Caption = a1
If Label4.Caption = "" Then
    a1 = InputBox("Please Enter the name of player # 2. ")
Label4.Caption = a2
If Label5.Caption = "" Then
    a1 = InputBox("Please Enter the name of player # 3. ")
Label5.Caption = a3
If Label6.Caption = "" Then
    a1 = InputBox("Please Enter the name of player # 4. ")
Label6.Caption = a4
End If
End If
End If
End If
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label9.Caption = Str(a)
Label10.Caption = Str(b)
Label11.Caption = Str(c)
Label12.Caption = Str(d)
End Sub

Private Sub roll1_Click()
Randomize
no1 = Int((6 * Rnd) + 1)
m1.Caption = Str(no1)
Select Case no1
    Case 1
    c1.Visible = False
    c2.Visible = False
    c3.Visible = False
    c4.Visible = False
    c5.Visible = False
    c6.Visible = False
    c7.Visible = True
    Case 2
    c1.Visible = True
    c2.Visible = False
    c3.Visible = False
    c4.Visible = False
    c5.Visible = False
    c6.Visible = True
    c7.Visible = False
    Case 3
    c1.Visible = True
    c2.Visible = False
    c3.Visible = False
    c4.Visible = False
    c5.Visible = False
    c6.Visible = True
    c7.Visible = True
    Case 4
    c1.Visible = True
    c2.Visible = False
    c3.Visible = True
    c4.Visible = True
    c5.Visible = False
    c6.Visible = True
    c7.Visible = False
    Case 5
    c1.Visible = True
    c2.Visible = False
    c3.Visible = True
    c4.Visible = True
    c5.Visible = False
    c6.Visible = True
    c7.Visible = True
    Case 6
    c1.Visible = True
    c2.Visible = True
    c3.Visible = True
    c4.Visible = True
    c5.Visible = True
    c6.Visible = True
    c7.Visible = False
End Select
If a + no1 > 100 Then
    a = a - no1
    s1(a).Visible = True
ElseIf a + no1 = 100 Then
    roll3.Enabled = False
    roll1.Enabled = False
    roll2.Enabled = False
    roll4.Enabled = False
    MsgBox "Yipee..Player 1 wins!!!"
End If
a = a + no1
s1(a).Visible = True
s1(a - no1).Visible = False
If a = 45 Then
    MsgBox "Opps!! You landed on a snake. Go back to 18."
    s1(a).Visible = False
    a = 18
    s1(a).Visible = True
ElseIf a = 99 Then
    MsgBox "Opps!! You landed on a snake. Go back to 8."
    s1(a).Visible = False
    a = 8
    s1(a).Visible = True
ElseIf a = 59 Then
    MsgBox "Opps!! You landed on a snake. Go back to 21."
    s1(a).Visible = False
    a = 21
    s1(a).Visible = True
ElseIf a = 69 Then
    MsgBox "Opps!! You landed on a snake. Go back to 25"
    s1(a).Visible = False
    a = 25
    s1(a).Visible = True
ElseIf a = 75 Then
    MsgBox "Opps!! You landed on a snake. Go back to 41."
    s1(a).Visible = False
    a = 41
    s1(a).Visible = True
End If
If a = 12 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 72."
    s1(12).Visible = False
    s1(72).Visible = True
    a = 72
ElseIf a = 44 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 82."
    s1(a).Visible = False
    a = 82
    s1(a).Visible = True
    
ElseIf a = 7 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 35."
    s1(a).Visible = False
    a = 35
    s1(a).Visible = True
ElseIf a = 87 Then
    MsgBox "Opps!! You landed on a snake. Go back to 32"
    s1(a).Visible = False
    a = 32
    s1(a).Visible = True
End If
If a = 4 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 4 is a square of which number. ")
    End If
If t1 = "2" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s1(4).Visible = False
    s1(6).Visible = True
    a = 6
Else
End If
End If
If a = 9 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 9 is a square of which number. ")
    End If
If t1 = "3" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s1(9).Visible = False
    s1(11).Visible = True
    a = 11
Else
End If
End If
If a = 16 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 16 is a square of which number. ")
    End If
If t1 = "4" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s1(16).Visible = False
    s1(19).Visible = True
    a = 19
Else
End If
End If
If a = 25 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 25 is a square of which number. ")
    End If
If t1 = "5" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s1(25).Visible = False
    s1(27).Visible = True
    a = 27
Else
End If
End If
If a = 36 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 36 is a square of which number. ")
    End If
If t1 = "6" Then
    MsgBox ("Yipee.. there is a +4 for you")
    s1(36).Visible = False
    s1(40).Visible = True
    a = 40
Else
    MsgBox ("Oh! no wrong answer.")
End If
End If
If a = 49 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 49 is a square of which number. ")
    End If
If t1 = "7" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s1(49).Visible = False
    s1(52).Visible = True
    a = 52
Else
End If
End If
If a = 64 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 64 is a square of which number. ")
    End If
If t1 = "8" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s1(64).Visible = False
    s1(66).Visible = True
    a = 66
Else
End If
End If
If a = 81 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 81 is a square of which number. ")
    End If
If t1 = "9" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s1(81).Visible = False
    s1(84).Visible = True
    a = 84
Else
End If
End If
Label9.Caption = Str(a)
If p4.Value = 1 Then
roll1.Visible = False
roll2.Visible = True
roll3.Visible = False
roll4.Visible = False
End If
If p3.Value = 1 Then
roll1.Visible = False
roll2.Visible = True
roll3.Visible = False
End If
If p2.Value = 1 Then
roll1.Visible = False
roll2.Visible = True
End If
End Sub

Private Sub roll2_Click()
Randomize
no2 = Int((6 * Rnd) + 1)
m2.Caption = Str(no2)
Select Case no2
    Case 1
    c11.Visible = False
    c12.Visible = False
    c13.Visible = False
    c14.Visible = False
    c15.Visible = False
    c16.Visible = False
    c17.Visible = True
    Case 2
    c11.Visible = True
    c12.Visible = False
    c13.Visible = False
    c14.Visible = False
    c15.Visible = False
    c16.Visible = True
    c17.Visible = False
    Case 3
    c11.Visible = True
    c12.Visible = False
    c13.Visible = False
    c14.Visible = False
    c15.Visible = False
    c16.Visible = True
    c17.Visible = True
    Case 4
    c11.Visible = True
    c12.Visible = False
    c13.Visible = True
    c14.Visible = True
    c15.Visible = False
    c16.Visible = True
    c17.Visible = False
    Case 5
    c11.Visible = True
    c12.Visible = False
    c13.Visible = True
    c14.Visible = True
    c15.Visible = False
    c16.Visible = True
    c17.Visible = True
    Case 6
    c11.Visible = True
    c12.Visible = True
    c13.Visible = True
    c14.Visible = True
    c15.Visible = True
    c16.Visible = True
    c17.Visible = False
End Select
If b + no2 > 100 Then
    b = b - no2
    s2(b).Visible = True
ElseIf b + no2 = 100 Then
    roll3.Enabled = False
    roll1.Enabled = False
    roll2.Enabled = False
    roll4.Enabled = False
    MsgBox "Yipee..Player 2 Wins!!!"
End If
b = b + no2
s2(b).Visible = True
s2(b - no2).Visible = False
If b = 45 Then
    MsgBox "Opps!! You landed on a snake. Go back to 18."
    s2(b).Visible = False
    b = 18
    s2(b).Visible = True
ElseIf b = 99 Then
    MsgBox "Opps!! You landed on a snake. Go back to 8."
    s2(b).Visible = False
    b = 8
    s2(b).Visible = True
ElseIf b = 59 Then
    MsgBox "Opps!! You landed on a snake. Go back to 21."
    s2(b).Visible = False
    b = 21
    s2(b).Visible = True
ElseIf b = 69 Then
    MsgBox "Opps!! You landed on a snake. Go back to 25"
    s2(b).Visible = False
    b = 25
    s2(b).Visible = True
ElseIf b = 75 Then
    MsgBox "Opps!! You landed on a snake. Go back to 41."
    s2(b).Visible = False
    b = 41
    s2(b).Visible = True
End If
If b = 12 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 72."
    s2(b).Visible = False
    b = 72
    s2(b).Visible = True
ElseIf b = 44 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 82."
    s2(b).Visible = False
    b = 82
    s2(b).Visible = True
ElseIf b = 7 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 35."
    s2(b).Visible = False
    b = 35
    s2(b).Visible = True
ElseIf b = 87 Then
    MsgBox "Opps!! You landed on a snake. Go back to 32"
    s2(b).Visible = False
    b = 32
    s2(b).Visible = True
End If
If b = 4 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 4 is a square of which number. ")
    End If
If t1 = "2" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s2(4).Visible = False
    s2(6).Visible = True
    b = 6
Else
End If
End If
If b = 9 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 9 is a square of which number. ")
    End If
If t1 = "3" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s2(9).Visible = False
    s2(11).Visible = True
    b = 11
Else
End If
End If
If b = 16 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 16 is a square of which number. ")
    End If
If t1 = "4" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s2(16).Visible = False
    s2(19).Visible = True
    b = 19
Else
End If
End If
If b = 25 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 25 is a square of which number. ")
    End If
If t1 = "5" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s2(25).Visible = False
    s2(27).Visible = True
    b = 27
Else
End If
End If
If b = 36 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 36 is a square of which number. ")
    End If
If t1 = "6" Then
    MsgBox ("Yipee.. there is a +4 for you")
    s2(36).Visible = False
    s2(40).Visible = True
    b = 40
End If
End If
If b = 49 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 49 is a square of which number. ")
    End If
If t1 = "7" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s2(49).Visible = False
    s2(52).Visible = True
    b = 52
End If
End If
If b = 64 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 64 is a square of which number. ")
    End If
If t1 = "8" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s2(64).Visible = False
    s2(66).Visible = True
    b = 66
End If
End If
If b = 81 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 81 is a square of which number. ")
    End If
If t1 = "9" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s2(81).Visible = False
    s2(84).Visible = True
    b = 84
End If
End If
Label9.Caption = Str(a)
Label10.Caption = Str(b)
If p4.Value = 1 Then
roll3.Visible = True
roll2.Visible = False
roll4.Visible = False
roll1.Visible = False
End If
If p3.Value = 1 Then
roll2.Visible = False
roll3.Visible = True
roll1.Visible = False
End If
If p2.Value = 1 Then
roll1.Visible = True
roll2.Visible = False
End If
End Sub

Private Sub roll3_Click()
Randomize
no3 = Int((6 * Rnd) + 1)
m3.Caption = Str(no3)
Select Case no3
    Case 1
    c21.Visible = False
    c22.Visible = False
    c23.Visible = False
    c24.Visible = False
    c25.Visible = False
    c26.Visible = False
    c27.Visible = True
    Case 2
    c21.Visible = True
    c22.Visible = False
    c23.Visible = False
    c24.Visible = False
    c25.Visible = False
    c26.Visible = True
    c27.Visible = False
    Case 3
    c21.Visible = True
    c22.Visible = False
    c23.Visible = False
    c24.Visible = False
    c25.Visible = False
    c26.Visible = True
    c27.Visible = True
    Case 4
    c21.Visible = True
    c22.Visible = False
    c23.Visible = True
    c24.Visible = True
    c25.Visible = False
    c26.Visible = True
    c27.Visible = False
    Case 5
    c21.Visible = True
    c22.Visible = False
    c23.Visible = True
    c24.Visible = True
    c25.Visible = False
    c26.Visible = True
    c27.Visible = True
    Case 6
    c21.Visible = True
    c22.Visible = True
    c23.Visible = True
    c24.Visible = True
    c25.Visible = True
    c26.Visible = True
    c27.Visible = False
End Select
If c + no3 > 100 Then
    c = c - no3
    s3(c).Visible = True
ElseIf c + no3 = 100 Then
    roll3.Enabled = False
    roll1.Enabled = False
    roll2.Enabled = False
    roll4.Enabled = False
    MsgBox "Yipee..Player 3 Wins!!!"
End If
c = c + no3
s3(c).Visible = True
s3(c - no3).Visible = False
If c = 45 Then
    MsgBox "Opps!! You landed on a snake. Go back to 18."
    s3(c).Visible = False
    c = 18
    s3(c).Visible = True
ElseIf c = 99 Then
    MsgBox "Opps!! You landed on a snake. Go back to 8."
    s3(c).Visible = False
    c = 8
    s3(c).Visible = True
ElseIf c = 59 Then
    MsgBox "Opps!! You landed on a snake. Go back to 21."
    s3(c).Visible = False
    c = 21
    s3(c).Visible = True
ElseIf c = 69 Then
    MsgBox "Opps!! You landed on a snake. Go back to 25"
    s3(c).Visible = False
    c = 25
    s3(c).Visible = True
ElseIf c = 75 Then
    MsgBox "Opps!! You landed on a snake. Go back to 41."
    s3(c).Visible = False
    c = 41
    s3(c).Visible = True
End If
If c = 12 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 72."
    s3(c).Visible = False
    c = 72
    s3(c).Visible = True
ElseIf c = 44 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 82."
    s3(c).Visible = False
    c = 82
    s3(c).Visible = True
ElseIf c = 7 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 35."
    s3(c).Visible = False
    c = 35
    s3(c).Visible = True
ElseIf c = 87 Then
    MsgBox "Opps!! You landed on a snake. Go back to 32"
    s3(c).Visible = False
    c = 32
    s3(c).Visible = True
End If
If c = 4 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 4 is a square of which number. ")
    End If
If t1 = "2" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s3(4).Visible = False
    s3(6).Visible = True
    c = 6
End If
End If
If c = 9 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 9 is a square of which number. ")
    End If
If t1 = "3" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s3(9).Visible = False
    s3(11).Visible = True
    c = 11
End If
End If
If c = 16 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 16 is a square of which number. ")
    End If
If t1 = "4" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s3(16).Visible = False
    s3(19).Visible = True
    c = 19
End If
End If
If c = 25 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 25 is a square of which number. ")
    End If
If t1 = "5" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s3(25).Visible = False
    s3(27).Visible = True
    c = 27
End If
End If
If c = 36 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 36 is a square of which number. ")
    End If
If t1 = "6" Then
    MsgBox ("Yipee.. there is a +4 for you")
    s3(36).Visible = False
    s3(40).Visible = True
    c = 40
End If
End If
If c = 49 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 49 is a square of which number. ")
    End If
If t1 = "7" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s3(49).Visible = False
    s3(52).Visible = True
    c = 52
End If
End If
If c = 64 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 64 is a square of which number. ")
    End If
If t1 = "8" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s3(64).Visible = False
    s3(66).Visible = True
    c = 66
End If
End If
If c = 81 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 81 is a square of which number. ")
    End If
If t1 = "9" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s3(81).Visible = False
    s3(84).Visible = True
    c = 84
End If
End If
Label9.Caption = Str(a)
Label10.Caption = Str(b)
Label11.Caption = Str(c)
If p4.Value = 1 Then
roll4.Visible = True
roll2.Visible = False
roll3.Visible = False
roll1.Visible = False
End If
If p3.Value = 1 Then
roll3.Visible = False
roll1.Visible = True
roll2.Visible = False
End If
End Sub

Private Sub roll4_Click()
Randomize
no4 = Int((6 * Rnd) + 1)
m4.Caption = Str(no4)
Select Case no4
    Case 1
    c31.Visible = False
    c32.Visible = False
    c33.Visible = False
    c34.Visible = False
    c35.Visible = False
    c36.Visible = False
    c37.Visible = True
    Case 2
    c31.Visible = True
    c32.Visible = False
    c33.Visible = False
    c34.Visible = False
    c35.Visible = False
    c36.Visible = True
    c37.Visible = False
    Case 3
    c31.Visible = True
    c32.Visible = False
    c33.Visible = False
    c34.Visible = False
    c35.Visible = False
    c36.Visible = True
    c37.Visible = True
    Case 4
    c31.Visible = True
    c32.Visible = False
    c33.Visible = True
    c34.Visible = True
    c35.Visible = False
    c36.Visible = True
    c37.Visible = False
    Case 5
    c31.Visible = True
    c32.Visible = False
    c33.Visible = True
    c34.Visible = True
    c35.Visible = False
    c36.Visible = True
    c37.Visible = True
    Case 6
    c31.Visible = True
    c32.Visible = True
    c33.Visible = True
    c34.Visible = True
    c35.Visible = True
    c36.Visible = True
    c37.Visible = False
End Select
If d + no4 > 100 Then
    d = d - no4
    s4(d).Visible = True
ElseIf d + no4 = 100 Then
    roll3.Enabled = False
    roll1.Enabled = False
    roll2.Enabled = False
    roll4.Enabled = False
    MsgBox "Yipee..Player 4 Wins!!!"
End If
d = d + no4
s4(d).Visible = True
s4(d - no4).Visible = False
If d = 45 Then
    MsgBox "Opps!! You landed on a snake. Go back to 18."
    s4(d).Visible = False
    d = 18
    s4(d).Visible = True
ElseIf d = 99 Then
    MsgBox "Opps!! You landed on a snake. Go back to 8."
    s4(d).Visible = False
    d = 8
    s4(d).Visible = True
ElseIf d = 59 Then
    MsgBox "Opps!! You landed on a snake. Go back to 21."
    s4(d).Visible = False
    d = 21
    s4(d).Visible = True
ElseIf d = 69 Then
    MsgBox "Opps!! You landed on a snake. Go back to 25"
    s4(d).Visible = False
    d = 25
    s4(d).Visible = True
ElseIf d = 75 Then
    MsgBox "Opps!! You landed on a snake. Go back to 41."
    s4(d).Visible = False
    d = 41
    s4(d).Visible = True
End If
If d = 12 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 72."
    s4(d).Visible = False
    d = 72
    s4(d).Visible = True
ElseIf d = 44 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 82."
    s4(d).Visible = False
    d = 82
    s4(d).Visible = True
ElseIf d = 7 Then
    MsgBox "Yippee!! There is a ladder for you. Go to 35."
    s4(d).Visible = False
    d = 35
    s4(d).Visible = True
ElseIf d = 87 Then
    MsgBox "Opps!! You landed on a snake. Go back to 32"
    s4(d).Visible = False
    d = 32
    s4(d).Visible = True
End If
If d = 2 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 4 is a square of which number. ")
    End If
If t1 = "2" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s4(9).Visible = False
    s4(11).Visible = True
    d = 11
End If
End If
If d = 9 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 9 is a square of which number. ")
    End If
If t1 = "3" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s4(9).Visible = False
    s4(11).Visible = True
    d = 11
End If
End If
If d = 16 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 16 is a square of which number. ")
    End If
If t1 = "4" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s4(16).Visible = False
    s4(19).Visible = True
    d = 19
End If
End If
If d = 25 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 25 is a square of which number. ")
    End If
If t1 = "5" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s4(25).Visible = False
    s4(27).Visible = True
    d = 27
End If
End If
If d = 36 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 36 is a square of which number. ")
    End If
If t1 = "6" Then
    MsgBox ("Yipee.. there is a +4 for you")
    s4(36).Visible = False
    s4(40).Visible = True
    d = 40
End If
End If
If d = 49 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 49 is a square of whidh number. ")
    End If
If t1 = "7" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s4(49).Visible = False
    s4(52).Visible = True
    d = 52
End If
End If
If d = 64 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 64 is a square of which number. ")
    End If
If t1 = "8" Then
    MsgBox ("Yipee.. there is a +2 for you")
    s4(64).Visible = False
    s4(66).Visible = True
    d = 66
End If
End If
If d = 81 Then
    ax = MsgBox("Do you want to answer a bonus question", vbYesNo)
    If ax = vbYes Then
        t1 = InputBox(" 81 is a square of which number. ")
    End If
If t1 = "9" Then
    MsgBox ("Yipee.. there is a +3 for you")
    s4(81).Visible = False
    s4(84).Visible = True
    d = 84
End If
End If
Label9.Caption = Str(a)
Label10.Caption = Str(b)
Label11.Caption = Str(c)
Label12.Caption = Str(d)
If p4.Value = 1 Then
roll4.Visible = False
roll1.Visible = True
roll2.Visible = False
roll3.Visible = False
End If
End Sub


