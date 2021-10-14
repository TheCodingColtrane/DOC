VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCarregandoGrande 
   Caption         =   "Carregando..."
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10380
   OleObjectBlob   =   "FrmCarregandoGrande.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCarregandoGrande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
Dim Agencia, Conta, Usuario, Senha As String
Dim AgenciaObjJS, ContaObjJS, MenuDistribuicaoRaizObjJS, MenuDistribuicaoLinhaObjJS, MenuDistribuicaoObjJS  As Variant
Dim DOM As MSHTML.HTMLDocument

IE.Navigate "https://www.bancodoc.com.br/Home/index.aspx"

Do While IE.ReadyState <> 4: DoEvents: Loop
Set DOM = IE.Document
Set AgenciaObjJS = DOM.getElementById("ctl00_txtAgencia")
Set ContaObjJS = DOM.getElementById("ctl00_txtConta")
If Not AgenciaObjJS Is Nothing And Not ContaObjJS Is Nothing Then
Agencia = InputBox("Digite sua agência", "Login Bancodoc")
Conta = InputBox("Digite sua conta", "Login Bancodoc")
DOM.getElementById("ctl00_txtAgencia").Value = Agencia
DOM.getElementById("ctl00_txtConta").Value = Conta
DOM.getElementById("ctl00_btnLogar").Click
Usuario = InputBox("Digite seu Usuário", "Login Bancodoc")
Senha = InputBox("Digite sua Senha", "Login Bancodoc")
DOM.getElementById("ctl00_ContentPlaceHolder1_txtUsuario").Value = Usuario
DOM.getElementById("ctl00_ContentPlaceHolder1_txtSenha").Value = Senha
DOM.getElementById("ctl00_ContentPlaceHolder1_Enviar").Click
Set MenuDistribuicaoRaizObjJS = DOM.getElementById("mnuMenuTopon13Items").Style = "visibility: visible; display: inline; top: 225px; height: 20px; clip: rect(auto, auto, auto, auto); left: 88px; z-index: 1;"
DOM.getElementById("mnuMenuTopon13Items").onmouseover = "Menu_HoverDynamic(this)"
Else
DOM.getElementById("ctl00_btnEntrarSistema").Click
End If


End Sub


Private Sub WBBGIFCarregando_StatusTextChange(ByVal Text As String)

End Sub
