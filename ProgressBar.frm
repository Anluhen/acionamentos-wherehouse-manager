VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "UserForm1"
   ClientHeight    =   2988
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7944
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------- frmProgress Code Module ---------
Option Explicit

' Call this to move the bar and show a new status.
Public Sub UpdateProgress(ByVal percent As Long)
    Dim frameWidth As Single
    
    ' Ensure percent is between 0 and 100
    If percent < 0 Then percent = 0
    If percent > 100 Then percent = 100
    
    ' Display the percentage as text on the bar:
    Me.lblBar.Caption = percent & "%"
    
    ' Update the status label
    If percent >= 25 Then
        Me.Label1.Caption = "ChrW(&H2713) Informações extraídas da ZTMM091"
    ElseIf percent >= 50 Then
        Me.Label2.Caption = "ChrW(&H2713) Informações extraídas da VL10G"""
    ElseIf percent >= 75 Then
        Me.Label3.Caption = "ChrW(&H2713) Informações extraídas do Analysis"
    ElseIf percent >= 99 Then
        Me.Label4.Caption = "ChrW(&H2713) Datas destacadas"
    ElseIf percent >= 100 Then
        Me.lblBar.Caption = "CONCLUÍDO"
    End If
    
    ' Calculate new width for lblBar
    frameWidth = Me.FrameProgress.Width
    Me.lblBar.Width = (percent / 100) * frameWidth
    
    ' Force repaint
    DoEvents
End Sub

' Clear or reset form when it initializes (optional)
Private Sub UserForm_Initialize()
    Me.lblStatus.Caption = ""
    Me.lblBar.Width = 0
    
    Me.Label1.Caption = "Extraindo informações da ZTMM091"
    Me.Label2.Caption = "Extraindo informações da VL10G"""
    Me.Label3.Caption = "Extraindo informações do Analysis"
    Me.Label4.Caption = "Destacando datas"
    
End Sub
'--------------------------------------------

