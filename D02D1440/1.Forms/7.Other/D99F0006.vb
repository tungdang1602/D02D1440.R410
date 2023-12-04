'#------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong class này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Title: D99F0006
'# CreateUser: NGUYEN THI MINH HOA
'# CreateDate: 23/03/2007
'# ModifiedUser: NGUYEN THI MINH HOA
'# ModifiedDate:  16/05/2011
'# Bổ sung: không kiểm tra giới hạn dữ liệu khi nhập chuỗi: bỏ kiểm tra MaxLength
'# Bổ sung hàm ListFindClientServer
'# Bổ sung đối số FindServer
'# Bổ sung hàm ListFindClientServer và ListFindClient
'# Description:  Bổ sung tìm kiếm KHÔNG CHỨA; Bổ sung Tìm kiếm giá trị NULL kiểu chuỗi và số Client
'# Tìm theo tập hợp.  Sửa lỗi khi tìm kiếm có chứa chữ N. Bổ sung load lại giá trị gần nhất tìm kiếm trước đó (10 giá trị gần nhất)
'#------------------------------------------------------

Imports System
Imports C1.Win.C1List

Friend Class D99F0006 'Form Chứa Code Finder
    Inherits System.Windows.Forms.Form

    Public Structure XArray
        Public Col1 As Object
        Public Col2 As Object
        Public Col3 As Object
        Public Col4 As Object
        Public Col5 As Object

        Public Sub New(ByVal Col1Value As Object, ByVal Col2Value As Object, ByVal Col3Value As Object, ByVal Col4Value As Object, ByVal Col5Value As Object)
            Col1 = Col1Value
            Col2 = Col2Value
            Col3 = Col3Value
            Col4 = Col4Value
            Col5 = Col5Value
        End Sub

        Public ReadOnly Property getCol1() As Object
            Get
                Return Col1
            End Get
        End Property

        Public ReadOnly Property getCol2() As Object
            Get
                Return Col2
            End Get
        End Property

        Public ReadOnly Property getCol3() As Object
            Get
                Return Col3
            End Get
        End Property

        Public ReadOnly Property getCol4() As Object
            Get
                Return Col4
            End Get
        End Property

        Public ReadOnly Property getCol5() As Object
            Get
                Return Col5
            End Get
        End Property

    End Structure


    Friend WithEvents chk01 As System.Windows.Forms.CheckBox
    Friend WithEvents chk02 As System.Windows.Forms.CheckBox
    Friend WithEvents chk10 As System.Windows.Forms.CheckBox
    Friend WithEvents chk09 As System.Windows.Forms.CheckBox
    Friend WithEvents chk08 As System.Windows.Forms.CheckBox
    Friend WithEvents chk07 As System.Windows.Forms.CheckBox
    Friend WithEvents chk06 As System.Windows.Forms.CheckBox
    Friend WithEvents chk05 As System.Windows.Forms.CheckBox
    Friend WithEvents chk04 As System.Windows.Forms.CheckBox
    Friend WithEvents chk03 As System.Windows.Forms.CheckBox

    Private _useUnicode As Boolean
    Public WriteOnly Property UseUnicode() As Boolean
        Set(ByVal vNewValue As Boolean)
            _useUnicode = vNewValue
        End Set
    End Property

    Private _formID As String
    Public WriteOnly Property FormID() As String
        Set(ByVal Value As String)
            _formID = Value
        End Set
    End Property

    Private _mode As String
    Public WriteOnly Property Mode() As String
        Set(ByVal Value As String)
            _mode = Value
        End Set
    End Property

    Private _findServer As Boolean
    Public WriteOnly Property FindServer() As Boolean
        Set(ByVal Value As Boolean)
            _findServer = Value
        End Set
    End Property

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ButClose As System.Windows.Forms.Button
    Friend WithEvents ButFind As System.Windows.Forms.Button
    Friend WithEvents ButAnd As System.Windows.Forms.Button
    Friend WithEvents ButOr As System.Windows.Forms.Button
    Friend WithEvents ButBack As System.Windows.Forms.Button
    Friend WithEvents CboOperator01 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField01 As C1.Win.C1List.C1Combo
    Friend WithEvents TxtValueB01 As System.Windows.Forms.TextBox
    Friend WithEvents CboField02 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField03 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField04 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField05 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField06 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField07 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField08 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField09 As C1.Win.C1List.C1Combo
    Friend WithEvents CboField10 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator02 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator03 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator04 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator05 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator06 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator07 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator08 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator09 As C1.Win.C1List.C1Combo
    Friend WithEvents CboOperator10 As C1.Win.C1List.C1Combo
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtValueB02 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueB03 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueB04 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueB05 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueB06 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueB07 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueB08 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueB09 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueB10 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA02 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA03 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA04 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA05 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA06 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA07 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA08 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA09 As System.Windows.Forms.TextBox
    Friend WithEvents TxtValueA10 As System.Windows.Forms.TextBox
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Private WithEvents c1dateB01 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA01 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA03 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA02 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA10 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA09 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA08 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA07 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA06 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA05 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateA04 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB10 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB08 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB07 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB06 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB05 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB04 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB03 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB02 As C1.Win.C1Input.C1DateEdit
    Private WithEvents c1dateB09 As C1.Win.C1Input.C1DateEdit
    Friend WithEvents btnBackFirst As System.Windows.Forms.Button
    Friend WithEvents btnPrevious As System.Windows.Forms.Button
    Friend WithEvents TxtValueA01 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Style1 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D99F0006))
        Dim Style2 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style3 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style4 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style5 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style6 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style7 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style8 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style9 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style10 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style11 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style12 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style13 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style14 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style15 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style16 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style17 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style18 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style19 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style20 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style21 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style22 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style23 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style24 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style25 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style26 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style27 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style28 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style29 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style30 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style31 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style32 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style33 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style34 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style35 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style36 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style37 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style38 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style39 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style40 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style41 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style42 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style43 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style44 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style45 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style46 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style47 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style48 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style49 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style50 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style51 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style52 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style53 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style54 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style55 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style56 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style57 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style58 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style59 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style60 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style61 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style62 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style63 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style64 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style65 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style66 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style67 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style68 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style69 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style70 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style71 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style72 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style73 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style74 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style75 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style76 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style77 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style78 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style79 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style80 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style81 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style82 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style83 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style84 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style85 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style86 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style87 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style88 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style89 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style90 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style91 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style92 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style93 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style94 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style95 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style96 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style97 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style98 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style99 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style100 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style101 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style102 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style103 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style104 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style105 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style106 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style107 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style108 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style109 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style110 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style111 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style112 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style113 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style114 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style115 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style116 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style117 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style118 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style119 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style120 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style121 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style122 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style123 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style124 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style125 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style126 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style127 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style128 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style129 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style130 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style131 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style132 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style133 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style134 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style135 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style136 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style137 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style138 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style139 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style140 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style141 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style142 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style143 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style144 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style145 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style146 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style147 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style148 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style149 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style150 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style151 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style152 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style153 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style154 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style155 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style156 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style157 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style158 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style159 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Dim Style160 As C1.Win.C1List.Style = New C1.Win.C1List.Style
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chk10 = New System.Windows.Forms.CheckBox
        Me.chk09 = New System.Windows.Forms.CheckBox
        Me.chk08 = New System.Windows.Forms.CheckBox
        Me.chk07 = New System.Windows.Forms.CheckBox
        Me.chk06 = New System.Windows.Forms.CheckBox
        Me.chk05 = New System.Windows.Forms.CheckBox
        Me.chk04 = New System.Windows.Forms.CheckBox
        Me.chk03 = New System.Windows.Forms.CheckBox
        Me.chk02 = New System.Windows.Forms.CheckBox
        Me.chk01 = New System.Windows.Forms.CheckBox
        Me.c1dateB10 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateB08 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateB07 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateB06 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateB05 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateB04 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateB03 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateB02 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA10 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA09 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA08 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA07 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA06 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA05 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA04 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA03 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateA02 = New C1.Win.C1Input.C1DateEdit
        Me.TxtValueA01 = New System.Windows.Forms.TextBox
        Me.TxtValueA10 = New System.Windows.Forms.TextBox
        Me.TxtValueA09 = New System.Windows.Forms.TextBox
        Me.TxtValueA08 = New System.Windows.Forms.TextBox
        Me.TxtValueA07 = New System.Windows.Forms.TextBox
        Me.TxtValueA06 = New System.Windows.Forms.TextBox
        Me.TxtValueA05 = New System.Windows.Forms.TextBox
        Me.TxtValueA04 = New System.Windows.Forms.TextBox
        Me.TxtValueA03 = New System.Windows.Forms.TextBox
        Me.TxtValueA02 = New System.Windows.Forms.TextBox
        Me.TxtValueB10 = New System.Windows.Forms.TextBox
        Me.TxtValueB09 = New System.Windows.Forms.TextBox
        Me.TxtValueB08 = New System.Windows.Forms.TextBox
        Me.TxtValueB07 = New System.Windows.Forms.TextBox
        Me.TxtValueB06 = New System.Windows.Forms.TextBox
        Me.TxtValueB05 = New System.Windows.Forms.TextBox
        Me.TxtValueB04 = New System.Windows.Forms.TextBox
        Me.TxtValueB03 = New System.Windows.Forms.TextBox
        Me.TxtValueB02 = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.CboOperator10 = New C1.Win.C1List.C1Combo
        Me.CboOperator09 = New C1.Win.C1List.C1Combo
        Me.CboOperator08 = New C1.Win.C1List.C1Combo
        Me.CboOperator07 = New C1.Win.C1List.C1Combo
        Me.CboOperator06 = New C1.Win.C1List.C1Combo
        Me.CboOperator05 = New C1.Win.C1List.C1Combo
        Me.CboOperator04 = New C1.Win.C1List.C1Combo
        Me.CboOperator03 = New C1.Win.C1List.C1Combo
        Me.CboOperator02 = New C1.Win.C1List.C1Combo
        Me.CboField10 = New C1.Win.C1List.C1Combo
        Me.CboField09 = New C1.Win.C1List.C1Combo
        Me.CboField08 = New C1.Win.C1List.C1Combo
        Me.CboField07 = New C1.Win.C1List.C1Combo
        Me.CboField06 = New C1.Win.C1List.C1Combo
        Me.CboField05 = New C1.Win.C1List.C1Combo
        Me.CboField04 = New C1.Win.C1List.C1Combo
        Me.CboField03 = New C1.Win.C1List.C1Combo
        Me.CboField02 = New C1.Win.C1List.C1Combo
        Me.CboOperator01 = New C1.Win.C1List.C1Combo
        Me.CboField01 = New C1.Win.C1List.C1Combo
        Me.TxtValueB01 = New System.Windows.Forms.TextBox
        Me.c1dateA01 = New C1.Win.C1Input.C1DateEdit
        Me.c1dateB01 = New C1.Win.C1Input.C1DateEdit
        Me.ButAnd = New System.Windows.Forms.Button
        Me.ButClose = New System.Windows.Forms.Button
        Me.ButFind = New System.Windows.Forms.Button
        Me.ButOr = New System.Windows.Forms.Button
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.c1dateB09 = New C1.Win.C1Input.C1DateEdit
        Me.btnPrevious = New System.Windows.Forms.Button
        Me.btnBackFirst = New System.Windows.Forms.Button
        Me.ButBack = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        CType(Me.c1dateB10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB08, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB07, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB06, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB05, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB04, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB03, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB02, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA09, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA08, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA07, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA06, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA05, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA04, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA03, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA02, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator09, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator08, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator07, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator06, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator05, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator04, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator03, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator02, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField09, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField08, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField07, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField06, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField05, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField04, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField03, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField02, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboOperator01, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboField01, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateA01, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB01, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.c1dateB09, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.chk10)
        Me.GroupBox1.Controls.Add(Me.chk09)
        Me.GroupBox1.Controls.Add(Me.chk08)
        Me.GroupBox1.Controls.Add(Me.chk07)
        Me.GroupBox1.Controls.Add(Me.chk06)
        Me.GroupBox1.Controls.Add(Me.chk05)
        Me.GroupBox1.Controls.Add(Me.chk04)
        Me.GroupBox1.Controls.Add(Me.chk03)
        Me.GroupBox1.Controls.Add(Me.chk02)
        Me.GroupBox1.Controls.Add(Me.chk01)
        Me.GroupBox1.Controls.Add(Me.c1dateB10)
        Me.GroupBox1.Controls.Add(Me.c1dateB08)
        Me.GroupBox1.Controls.Add(Me.c1dateB07)
        Me.GroupBox1.Controls.Add(Me.c1dateB06)
        Me.GroupBox1.Controls.Add(Me.c1dateB05)
        Me.GroupBox1.Controls.Add(Me.c1dateB04)
        Me.GroupBox1.Controls.Add(Me.c1dateB03)
        Me.GroupBox1.Controls.Add(Me.c1dateB02)
        Me.GroupBox1.Controls.Add(Me.c1dateA10)
        Me.GroupBox1.Controls.Add(Me.c1dateA09)
        Me.GroupBox1.Controls.Add(Me.c1dateA08)
        Me.GroupBox1.Controls.Add(Me.c1dateA07)
        Me.GroupBox1.Controls.Add(Me.c1dateA06)
        Me.GroupBox1.Controls.Add(Me.c1dateA05)
        Me.GroupBox1.Controls.Add(Me.c1dateA04)
        Me.GroupBox1.Controls.Add(Me.c1dateA03)
        Me.GroupBox1.Controls.Add(Me.c1dateA02)
        Me.GroupBox1.Controls.Add(Me.TxtValueA01)
        Me.GroupBox1.Controls.Add(Me.TxtValueA10)
        Me.GroupBox1.Controls.Add(Me.TxtValueA09)
        Me.GroupBox1.Controls.Add(Me.TxtValueA08)
        Me.GroupBox1.Controls.Add(Me.TxtValueA07)
        Me.GroupBox1.Controls.Add(Me.TxtValueA06)
        Me.GroupBox1.Controls.Add(Me.TxtValueA05)
        Me.GroupBox1.Controls.Add(Me.TxtValueA04)
        Me.GroupBox1.Controls.Add(Me.TxtValueA03)
        Me.GroupBox1.Controls.Add(Me.TxtValueA02)
        Me.GroupBox1.Controls.Add(Me.TxtValueB10)
        Me.GroupBox1.Controls.Add(Me.TxtValueB09)
        Me.GroupBox1.Controls.Add(Me.TxtValueB08)
        Me.GroupBox1.Controls.Add(Me.TxtValueB07)
        Me.GroupBox1.Controls.Add(Me.TxtValueB06)
        Me.GroupBox1.Controls.Add(Me.TxtValueB05)
        Me.GroupBox1.Controls.Add(Me.TxtValueB04)
        Me.GroupBox1.Controls.Add(Me.TxtValueB03)
        Me.GroupBox1.Controls.Add(Me.TxtValueB02)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.CboOperator10)
        Me.GroupBox1.Controls.Add(Me.CboOperator09)
        Me.GroupBox1.Controls.Add(Me.CboOperator08)
        Me.GroupBox1.Controls.Add(Me.CboOperator07)
        Me.GroupBox1.Controls.Add(Me.CboOperator06)
        Me.GroupBox1.Controls.Add(Me.CboOperator05)
        Me.GroupBox1.Controls.Add(Me.CboOperator04)
        Me.GroupBox1.Controls.Add(Me.CboOperator03)
        Me.GroupBox1.Controls.Add(Me.CboOperator02)
        Me.GroupBox1.Controls.Add(Me.CboField10)
        Me.GroupBox1.Controls.Add(Me.CboField09)
        Me.GroupBox1.Controls.Add(Me.CboField08)
        Me.GroupBox1.Controls.Add(Me.CboField07)
        Me.GroupBox1.Controls.Add(Me.CboField06)
        Me.GroupBox1.Controls.Add(Me.CboField05)
        Me.GroupBox1.Controls.Add(Me.CboField04)
        Me.GroupBox1.Controls.Add(Me.CboField03)
        Me.GroupBox1.Controls.Add(Me.CboField02)
        Me.GroupBox1.Controls.Add(Me.CboOperator01)
        Me.GroupBox1.Controls.Add(Me.CboField01)
        Me.GroupBox1.Controls.Add(Me.TxtValueB01)
        Me.GroupBox1.Controls.Add(Me.c1dateA01)
        Me.GroupBox1.Controls.Add(Me.c1dateB01)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(6, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(528, 70)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Điều kiện tìm kiếm"
        '
        'chk10
        '
        Me.chk10.AutoSize = True
        Me.chk10.Location = New System.Drawing.Point(319, -44)
        Me.chk10.Name = "chk10"
        Me.chk10.Size = New System.Drawing.Size(15, 14)
        Me.chk10.TabIndex = 103
        Me.chk10.UseVisualStyleBackColor = True
        '
        'chk09
        '
        Me.chk09.AutoSize = True
        Me.chk09.Location = New System.Drawing.Point(319, -44)
        Me.chk09.Name = "chk09"
        Me.chk09.Size = New System.Drawing.Size(15, 14)
        Me.chk09.TabIndex = 102
        Me.chk09.UseVisualStyleBackColor = True
        '
        'chk08
        '
        Me.chk08.AutoSize = True
        Me.chk08.Location = New System.Drawing.Point(319, -44)
        Me.chk08.Name = "chk08"
        Me.chk08.Size = New System.Drawing.Size(15, 14)
        Me.chk08.TabIndex = 101
        Me.chk08.UseVisualStyleBackColor = True
        '
        'chk07
        '
        Me.chk07.AutoSize = True
        Me.chk07.Location = New System.Drawing.Point(319, -44)
        Me.chk07.Name = "chk07"
        Me.chk07.Size = New System.Drawing.Size(15, 14)
        Me.chk07.TabIndex = 100
        Me.chk07.UseVisualStyleBackColor = True
        '
        'chk06
        '
        Me.chk06.AutoSize = True
        Me.chk06.Location = New System.Drawing.Point(319, -44)
        Me.chk06.Name = "chk06"
        Me.chk06.Size = New System.Drawing.Size(15, 14)
        Me.chk06.TabIndex = 99
        Me.chk06.UseVisualStyleBackColor = True
        '
        'chk05
        '
        Me.chk05.AutoSize = True
        Me.chk05.Location = New System.Drawing.Point(319, -44)
        Me.chk05.Name = "chk05"
        Me.chk05.Size = New System.Drawing.Size(15, 14)
        Me.chk05.TabIndex = 98
        Me.chk05.UseVisualStyleBackColor = True
        '
        'chk04
        '
        Me.chk04.AutoSize = True
        Me.chk04.Location = New System.Drawing.Point(319, -44)
        Me.chk04.Name = "chk04"
        Me.chk04.Size = New System.Drawing.Size(15, 14)
        Me.chk04.TabIndex = 97
        Me.chk04.UseVisualStyleBackColor = True
        '
        'chk03
        '
        Me.chk03.AutoSize = True
        Me.chk03.Location = New System.Drawing.Point(319, -44)
        Me.chk03.Name = "chk03"
        Me.chk03.Size = New System.Drawing.Size(15, 14)
        Me.chk03.TabIndex = 96
        Me.chk03.UseVisualStyleBackColor = True
        '
        'chk02
        '
        Me.chk02.AutoSize = True
        Me.chk02.Location = New System.Drawing.Point(319, -44)
        Me.chk02.Name = "chk02"
        Me.chk02.Size = New System.Drawing.Size(15, 14)
        Me.chk02.TabIndex = 95
        Me.chk02.UseVisualStyleBackColor = True
        '
        'chk01
        '
        Me.chk01.AutoSize = True
        Me.chk01.Location = New System.Drawing.Point(319, 33)
        Me.chk01.Name = "chk01"
        Me.chk01.Size = New System.Drawing.Size(15, 14)
        Me.chk01.TabIndex = 94
        Me.chk01.UseVisualStyleBackColor = True
        '
        'c1dateB10
        '
        Me.c1dateB10.AutoSize = False
        Me.c1dateB10.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB10.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB10.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB10.Location = New System.Drawing.Point(305, -34)
        Me.c1dateB10.Name = "c1dateB10"
        Me.c1dateB10.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB10.TabIndex = 93
        Me.c1dateB10.Tag = Nothing
        Me.c1dateB10.TrimStart = True
        Me.c1dateB10.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateB08
        '
        Me.c1dateB08.AutoSize = False
        Me.c1dateB08.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB08.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB08.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB08.Location = New System.Drawing.Point(333, -34)
        Me.c1dateB08.Name = "c1dateB08"
        Me.c1dateB08.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB08.TabIndex = 92
        Me.c1dateB08.Tag = Nothing
        Me.c1dateB08.TrimStart = True
        Me.c1dateB08.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateB07
        '
        Me.c1dateB07.AutoSize = False
        Me.c1dateB07.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB07.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB07.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB07.Location = New System.Drawing.Point(317, -34)
        Me.c1dateB07.Name = "c1dateB07"
        Me.c1dateB07.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB07.TabIndex = 91
        Me.c1dateB07.Tag = Nothing
        Me.c1dateB07.TrimStart = True
        Me.c1dateB07.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateB06
        '
        Me.c1dateB06.AutoSize = False
        Me.c1dateB06.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB06.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB06.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB06.Location = New System.Drawing.Point(310, -34)
        Me.c1dateB06.Name = "c1dateB06"
        Me.c1dateB06.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB06.TabIndex = 90
        Me.c1dateB06.Tag = Nothing
        Me.c1dateB06.TrimStart = True
        Me.c1dateB06.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateB05
        '
        Me.c1dateB05.AutoSize = False
        Me.c1dateB05.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB05.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB05.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB05.Location = New System.Drawing.Point(336, -34)
        Me.c1dateB05.Name = "c1dateB05"
        Me.c1dateB05.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB05.TabIndex = 89
        Me.c1dateB05.Tag = Nothing
        Me.c1dateB05.TrimStart = True
        Me.c1dateB05.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateB04
        '
        Me.c1dateB04.AutoSize = False
        Me.c1dateB04.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB04.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB04.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB04.Location = New System.Drawing.Point(430, -34)
        Me.c1dateB04.Name = "c1dateB04"
        Me.c1dateB04.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB04.TabIndex = 88
        Me.c1dateB04.Tag = Nothing
        Me.c1dateB04.TrimStart = True
        Me.c1dateB04.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateB03
        '
        Me.c1dateB03.AutoSize = False
        Me.c1dateB03.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB03.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB03.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB03.Location = New System.Drawing.Point(309, -34)
        Me.c1dateB03.Name = "c1dateB03"
        Me.c1dateB03.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB03.TabIndex = 87
        Me.c1dateB03.Tag = Nothing
        Me.c1dateB03.TrimStart = True
        Me.c1dateB03.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateB02
        '
        Me.c1dateB02.AutoSize = False
        Me.c1dateB02.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB02.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB02.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB02.Location = New System.Drawing.Point(381, -34)
        Me.c1dateB02.Name = "c1dateB02"
        Me.c1dateB02.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB02.TabIndex = 86
        Me.c1dateB02.Tag = Nothing
        Me.c1dateB02.TrimStart = True
        Me.c1dateB02.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA10
        '
        Me.c1dateA10.AutoSize = False
        Me.c1dateA10.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA10.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA10.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA10.Location = New System.Drawing.Point(325, -36)
        Me.c1dateA10.Name = "c1dateA10"
        Me.c1dateA10.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA10.TabIndex = 85
        Me.c1dateA10.Tag = Nothing
        Me.c1dateA10.TrimStart = True
        Me.c1dateA10.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA09
        '
        Me.c1dateA09.AutoSize = False
        Me.c1dateA09.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA09.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA09.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA09.Location = New System.Drawing.Point(303, -36)
        Me.c1dateA09.Name = "c1dateA09"
        Me.c1dateA09.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA09.TabIndex = 84
        Me.c1dateA09.Tag = Nothing
        Me.c1dateA09.TrimStart = True
        Me.c1dateA09.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA08
        '
        Me.c1dateA08.AutoSize = False
        Me.c1dateA08.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA08.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA08.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA08.Location = New System.Drawing.Point(312, -36)
        Me.c1dateA08.Name = "c1dateA08"
        Me.c1dateA08.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA08.TabIndex = 83
        Me.c1dateA08.Tag = Nothing
        Me.c1dateA08.TrimStart = True
        Me.c1dateA08.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA07
        '
        Me.c1dateA07.AutoSize = False
        Me.c1dateA07.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA07.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA07.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA07.Location = New System.Drawing.Point(305, -36)
        Me.c1dateA07.Name = "c1dateA07"
        Me.c1dateA07.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA07.TabIndex = 82
        Me.c1dateA07.Tag = Nothing
        Me.c1dateA07.TrimStart = True
        Me.c1dateA07.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA06
        '
        Me.c1dateA06.AutoSize = False
        Me.c1dateA06.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA06.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA06.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA06.Location = New System.Drawing.Point(305, -36)
        Me.c1dateA06.Name = "c1dateA06"
        Me.c1dateA06.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA06.TabIndex = 81
        Me.c1dateA06.Tag = Nothing
        Me.c1dateA06.TrimStart = True
        Me.c1dateA06.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA05
        '
        Me.c1dateA05.AutoSize = False
        Me.c1dateA05.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA05.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA05.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA05.Location = New System.Drawing.Point(299, -36)
        Me.c1dateA05.Name = "c1dateA05"
        Me.c1dateA05.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA05.TabIndex = 80
        Me.c1dateA05.Tag = Nothing
        Me.c1dateA05.TrimStart = True
        Me.c1dateA05.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA04
        '
        Me.c1dateA04.AutoSize = False
        Me.c1dateA04.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA04.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA04.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA04.Location = New System.Drawing.Point(303, -36)
        Me.c1dateA04.Name = "c1dateA04"
        Me.c1dateA04.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA04.TabIndex = 79
        Me.c1dateA04.Tag = Nothing
        Me.c1dateA04.TrimStart = True
        Me.c1dateA04.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA03
        '
        Me.c1dateA03.AutoSize = False
        Me.c1dateA03.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA03.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA03.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA03.Location = New System.Drawing.Point(303, -36)
        Me.c1dateA03.Name = "c1dateA03"
        Me.c1dateA03.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA03.TabIndex = 78
        Me.c1dateA03.Tag = Nothing
        Me.c1dateA03.TrimStart = True
        Me.c1dateA03.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateA02
        '
        Me.c1dateA02.AutoSize = False
        Me.c1dateA02.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA02.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA02.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA02.Location = New System.Drawing.Point(292, -36)
        Me.c1dateA02.Name = "c1dateA02"
        Me.c1dateA02.Size = New System.Drawing.Size(100, 22)
        Me.c1dateA02.TabIndex = 77
        Me.c1dateA02.Tag = Nothing
        Me.c1dateA02.TrimStart = True
        Me.c1dateA02.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'TxtValueA01
        '
        Me.TxtValueA01.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA01.Location = New System.Drawing.Point(307, 29)
        Me.TxtValueA01.Name = "TxtValueA01"
        Me.TxtValueA01.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA01.TabIndex = 74
        '
        'TxtValueA10
        '
        Me.TxtValueA10.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA10.Location = New System.Drawing.Point(193, -28)
        Me.TxtValueA10.Name = "TxtValueA10"
        Me.TxtValueA10.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA10.TabIndex = 73
        '
        'TxtValueA09
        '
        Me.TxtValueA09.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA09.Location = New System.Drawing.Point(191, -30)
        Me.TxtValueA09.Name = "TxtValueA09"
        Me.TxtValueA09.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA09.TabIndex = 72
        '
        'TxtValueA08
        '
        Me.TxtValueA08.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA08.Location = New System.Drawing.Point(189, -32)
        Me.TxtValueA08.Name = "TxtValueA08"
        Me.TxtValueA08.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA08.TabIndex = 71
        '
        'TxtValueA07
        '
        Me.TxtValueA07.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA07.Location = New System.Drawing.Point(187, -34)
        Me.TxtValueA07.Name = "TxtValueA07"
        Me.TxtValueA07.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA07.TabIndex = 70
        '
        'TxtValueA06
        '
        Me.TxtValueA06.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA06.Location = New System.Drawing.Point(185, -36)
        Me.TxtValueA06.Name = "TxtValueA06"
        Me.TxtValueA06.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA06.TabIndex = 69
        '
        'TxtValueA05
        '
        Me.TxtValueA05.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA05.Location = New System.Drawing.Point(183, -38)
        Me.TxtValueA05.Name = "TxtValueA05"
        Me.TxtValueA05.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA05.TabIndex = 68
        '
        'TxtValueA04
        '
        Me.TxtValueA04.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA04.Location = New System.Drawing.Point(181, -40)
        Me.TxtValueA04.Name = "TxtValueA04"
        Me.TxtValueA04.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA04.TabIndex = 67
        '
        'TxtValueA03
        '
        Me.TxtValueA03.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA03.Location = New System.Drawing.Point(179, -42)
        Me.TxtValueA03.Name = "TxtValueA03"
        Me.TxtValueA03.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA03.TabIndex = 66
        '
        'TxtValueA02
        '
        Me.TxtValueA02.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueA02.Location = New System.Drawing.Point(177, -44)
        Me.TxtValueA02.Name = "TxtValueA02"
        Me.TxtValueA02.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueA02.TabIndex = 65
        '
        'TxtValueB10
        '
        Me.TxtValueB10.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB10.Location = New System.Drawing.Point(175, -46)
        Me.TxtValueB10.Name = "TxtValueB10"
        Me.TxtValueB10.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB10.TabIndex = 64
        '
        'TxtValueB09
        '
        Me.TxtValueB09.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB09.Location = New System.Drawing.Point(328, -24)
        Me.TxtValueB09.Name = "TxtValueB09"
        Me.TxtValueB09.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB09.TabIndex = 63
        '
        'TxtValueB08
        '
        Me.TxtValueB08.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB08.Location = New System.Drawing.Point(326, -26)
        Me.TxtValueB08.Name = "TxtValueB08"
        Me.TxtValueB08.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB08.TabIndex = 62
        '
        'TxtValueB07
        '
        Me.TxtValueB07.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB07.Location = New System.Drawing.Point(324, -26)
        Me.TxtValueB07.Name = "TxtValueB07"
        Me.TxtValueB07.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB07.TabIndex = 61
        '
        'TxtValueB06
        '
        Me.TxtValueB06.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB06.Location = New System.Drawing.Point(322, -28)
        Me.TxtValueB06.Name = "TxtValueB06"
        Me.TxtValueB06.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB06.TabIndex = 60
        '
        'TxtValueB05
        '
        Me.TxtValueB05.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB05.Location = New System.Drawing.Point(320, -30)
        Me.TxtValueB05.Name = "TxtValueB05"
        Me.TxtValueB05.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB05.TabIndex = 59
        '
        'TxtValueB04
        '
        Me.TxtValueB04.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB04.Location = New System.Drawing.Point(318, -32)
        Me.TxtValueB04.Name = "TxtValueB04"
        Me.TxtValueB04.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB04.TabIndex = 58
        '
        'TxtValueB03
        '
        Me.TxtValueB03.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB03.Location = New System.Drawing.Point(316, -34)
        Me.TxtValueB03.Name = "TxtValueB03"
        Me.TxtValueB03.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB03.TabIndex = 57
        '
        'TxtValueB02
        '
        Me.TxtValueB02.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB02.Location = New System.Drawing.Point(314, -36)
        Me.TxtValueB02.Name = "TxtValueB02"
        Me.TxtValueB02.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB02.TabIndex = 56
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(398, -18)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(45, 13)
        Me.Label10.TabIndex = 55
        Me.Label10.Text = "Label10"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(396, -20)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 54
        Me.Label9.Text = "Label9"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(394, -22)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(39, 13)
        Me.Label8.TabIndex = 53
        Me.Label8.Text = "Label8"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(392, -24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(39, 13)
        Me.Label7.TabIndex = 52
        Me.Label7.Text = "Label7"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(390, -26)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 13)
        Me.Label6.TabIndex = 51
        Me.Label6.Text = "Label6"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(388, -26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 50
        Me.Label5.Text = "Label5"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(386, -28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 13)
        Me.Label4.TabIndex = 49
        Me.Label4.Text = "Label4"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(384, -30)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 48
        Me.Label3.Text = "Label3"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(382, -32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 47
        Me.Label2.Text = "Label2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(11, -28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 46
        Me.Label1.Text = "Label1"
        '
        'CboOperator10
        '
        Me.CboOperator10.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator10.Caption = ""
        Me.CboOperator10.CaptionHeight = 17
        Me.CboOperator10.CaptionStyle = Style1
        Me.CboOperator10.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator10.ColumnCaptionHeight = 17
        Me.CboOperator10.ColumnFooterHeight = 17
        Me.CboOperator10.ContentHeight = 17
        Me.CboOperator10.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator10.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator10.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator10.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator10.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator10.EditorHeight = 17
        Me.CboOperator10.EvenRowStyle = Style2
        Me.CboOperator10.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator10.FooterStyle = Style3
        Me.CboOperator10.HeadingStyle = Style4
        Me.CboOperator10.HighLightRowStyle = Style5
        Me.CboOperator10.Images.Add(CType(resources.GetObject("CboOperator10.Images"), System.Drawing.Image))
        Me.CboOperator10.ItemHeight = 15
        Me.CboOperator10.Location = New System.Drawing.Point(178, -29)
        Me.CboOperator10.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator10.MaxDropDownItems = CType(5, Short)
        Me.CboOperator10.MaxLength = 32767
        Me.CboOperator10.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator10.Name = "CboOperator10"
        Me.CboOperator10.OddRowStyle = Style6
        Me.CboOperator10.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator10.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator10.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator10.SelectedStyle = Style7
        Me.CboOperator10.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator10.Style = Style8
        Me.CboOperator10.TabIndex = 27
        Me.CboOperator10.PropBag = resources.GetString("CboOperator10.PropBag")
        '
        'CboOperator09
        '
        Me.CboOperator09.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator09.Caption = ""
        Me.CboOperator09.CaptionHeight = 17
        Me.CboOperator09.CaptionStyle = Style9
        Me.CboOperator09.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator09.ColumnCaptionHeight = 17
        Me.CboOperator09.ColumnFooterHeight = 17
        Me.CboOperator09.ContentHeight = 17
        Me.CboOperator09.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator09.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator09.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator09.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator09.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator09.EditorHeight = 17
        Me.CboOperator09.EvenRowStyle = Style10
        Me.CboOperator09.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator09.FooterStyle = Style11
        Me.CboOperator09.HeadingStyle = Style12
        Me.CboOperator09.HighLightRowStyle = Style13
        Me.CboOperator09.Images.Add(CType(resources.GetObject("CboOperator09.Images"), System.Drawing.Image))
        Me.CboOperator09.ItemHeight = 15
        Me.CboOperator09.Location = New System.Drawing.Point(176, -31)
        Me.CboOperator09.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator09.MaxDropDownItems = CType(5, Short)
        Me.CboOperator09.MaxLength = 32767
        Me.CboOperator09.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator09.Name = "CboOperator09"
        Me.CboOperator09.OddRowStyle = Style14
        Me.CboOperator09.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator09.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator09.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator09.SelectedStyle = Style15
        Me.CboOperator09.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator09.Style = Style16
        Me.CboOperator09.TabIndex = 26
        Me.CboOperator09.PropBag = resources.GetString("CboOperator09.PropBag")
        '
        'CboOperator08
        '
        Me.CboOperator08.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator08.Caption = ""
        Me.CboOperator08.CaptionHeight = 17
        Me.CboOperator08.CaptionStyle = Style17
        Me.CboOperator08.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator08.ColumnCaptionHeight = 17
        Me.CboOperator08.ColumnFooterHeight = 17
        Me.CboOperator08.ContentHeight = 17
        Me.CboOperator08.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator08.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator08.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator08.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator08.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator08.EditorHeight = 17
        Me.CboOperator08.EvenRowStyle = Style18
        Me.CboOperator08.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator08.FooterStyle = Style19
        Me.CboOperator08.HeadingStyle = Style20
        Me.CboOperator08.HighLightRowStyle = Style21
        Me.CboOperator08.Images.Add(CType(resources.GetObject("CboOperator08.Images"), System.Drawing.Image))
        Me.CboOperator08.ItemHeight = 15
        Me.CboOperator08.Location = New System.Drawing.Point(174, -33)
        Me.CboOperator08.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator08.MaxDropDownItems = CType(5, Short)
        Me.CboOperator08.MaxLength = 32767
        Me.CboOperator08.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator08.Name = "CboOperator08"
        Me.CboOperator08.OddRowStyle = Style22
        Me.CboOperator08.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator08.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator08.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator08.SelectedStyle = Style23
        Me.CboOperator08.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator08.Style = Style24
        Me.CboOperator08.TabIndex = 25
        Me.CboOperator08.PropBag = resources.GetString("CboOperator08.PropBag")
        '
        'CboOperator07
        '
        Me.CboOperator07.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator07.Caption = ""
        Me.CboOperator07.CaptionHeight = 17
        Me.CboOperator07.CaptionStyle = Style25
        Me.CboOperator07.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator07.ColumnCaptionHeight = 17
        Me.CboOperator07.ColumnFooterHeight = 17
        Me.CboOperator07.ContentHeight = 17
        Me.CboOperator07.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator07.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator07.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator07.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator07.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator07.EditorHeight = 17
        Me.CboOperator07.EvenRowStyle = Style26
        Me.CboOperator07.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator07.FooterStyle = Style27
        Me.CboOperator07.HeadingStyle = Style28
        Me.CboOperator07.HighLightRowStyle = Style29
        Me.CboOperator07.Images.Add(CType(resources.GetObject("CboOperator07.Images"), System.Drawing.Image))
        Me.CboOperator07.ItemHeight = 15
        Me.CboOperator07.Location = New System.Drawing.Point(172, -35)
        Me.CboOperator07.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator07.MaxDropDownItems = CType(5, Short)
        Me.CboOperator07.MaxLength = 32767
        Me.CboOperator07.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator07.Name = "CboOperator07"
        Me.CboOperator07.OddRowStyle = Style30
        Me.CboOperator07.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator07.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator07.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator07.SelectedStyle = Style31
        Me.CboOperator07.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator07.Style = Style32
        Me.CboOperator07.TabIndex = 24
        Me.CboOperator07.PropBag = resources.GetString("CboOperator07.PropBag")
        '
        'CboOperator06
        '
        Me.CboOperator06.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator06.Caption = ""
        Me.CboOperator06.CaptionHeight = 17
        Me.CboOperator06.CaptionStyle = Style33
        Me.CboOperator06.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator06.ColumnCaptionHeight = 17
        Me.CboOperator06.ColumnFooterHeight = 17
        Me.CboOperator06.ContentHeight = 17
        Me.CboOperator06.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator06.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator06.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator06.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator06.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator06.EditorHeight = 17
        Me.CboOperator06.EvenRowStyle = Style34
        Me.CboOperator06.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator06.FooterStyle = Style35
        Me.CboOperator06.HeadingStyle = Style36
        Me.CboOperator06.HighLightRowStyle = Style37
        Me.CboOperator06.Images.Add(CType(resources.GetObject("CboOperator06.Images"), System.Drawing.Image))
        Me.CboOperator06.ItemHeight = 15
        Me.CboOperator06.Location = New System.Drawing.Point(170, -37)
        Me.CboOperator06.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator06.MaxDropDownItems = CType(5, Short)
        Me.CboOperator06.MaxLength = 32767
        Me.CboOperator06.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator06.Name = "CboOperator06"
        Me.CboOperator06.OddRowStyle = Style38
        Me.CboOperator06.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator06.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator06.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator06.SelectedStyle = Style39
        Me.CboOperator06.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator06.Style = Style40
        Me.CboOperator06.TabIndex = 23
        Me.CboOperator06.PropBag = resources.GetString("CboOperator06.PropBag")
        '
        'CboOperator05
        '
        Me.CboOperator05.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator05.Caption = ""
        Me.CboOperator05.CaptionHeight = 17
        Me.CboOperator05.CaptionStyle = Style41
        Me.CboOperator05.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator05.ColumnCaptionHeight = 17
        Me.CboOperator05.ColumnFooterHeight = 17
        Me.CboOperator05.ContentHeight = 17
        Me.CboOperator05.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator05.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator05.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator05.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator05.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator05.EditorHeight = 17
        Me.CboOperator05.EvenRowStyle = Style42
        Me.CboOperator05.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator05.FooterStyle = Style43
        Me.CboOperator05.HeadingStyle = Style44
        Me.CboOperator05.HighLightRowStyle = Style45
        Me.CboOperator05.Images.Add(CType(resources.GetObject("CboOperator05.Images"), System.Drawing.Image))
        Me.CboOperator05.ItemHeight = 15
        Me.CboOperator05.Location = New System.Drawing.Point(168, -39)
        Me.CboOperator05.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator05.MaxDropDownItems = CType(5, Short)
        Me.CboOperator05.MaxLength = 32767
        Me.CboOperator05.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator05.Name = "CboOperator05"
        Me.CboOperator05.OddRowStyle = Style46
        Me.CboOperator05.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator05.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator05.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator05.SelectedStyle = Style47
        Me.CboOperator05.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator05.Style = Style48
        Me.CboOperator05.TabIndex = 22
        Me.CboOperator05.PropBag = resources.GetString("CboOperator05.PropBag")
        '
        'CboOperator04
        '
        Me.CboOperator04.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator04.Caption = ""
        Me.CboOperator04.CaptionHeight = 17
        Me.CboOperator04.CaptionStyle = Style49
        Me.CboOperator04.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator04.ColumnCaptionHeight = 17
        Me.CboOperator04.ColumnFooterHeight = 17
        Me.CboOperator04.ContentHeight = 17
        Me.CboOperator04.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator04.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator04.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator04.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator04.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator04.EditorHeight = 17
        Me.CboOperator04.EvenRowStyle = Style50
        Me.CboOperator04.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator04.FooterStyle = Style51
        Me.CboOperator04.HeadingStyle = Style52
        Me.CboOperator04.HighLightRowStyle = Style53
        Me.CboOperator04.Images.Add(CType(resources.GetObject("CboOperator04.Images"), System.Drawing.Image))
        Me.CboOperator04.ItemHeight = 15
        Me.CboOperator04.Location = New System.Drawing.Point(166, -41)
        Me.CboOperator04.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator04.MaxDropDownItems = CType(5, Short)
        Me.CboOperator04.MaxLength = 32767
        Me.CboOperator04.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator04.Name = "CboOperator04"
        Me.CboOperator04.OddRowStyle = Style54
        Me.CboOperator04.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator04.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator04.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator04.SelectedStyle = Style55
        Me.CboOperator04.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator04.Style = Style56
        Me.CboOperator04.TabIndex = 21
        Me.CboOperator04.PropBag = resources.GetString("CboOperator04.PropBag")
        '
        'CboOperator03
        '
        Me.CboOperator03.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator03.Caption = ""
        Me.CboOperator03.CaptionHeight = 17
        Me.CboOperator03.CaptionStyle = Style57
        Me.CboOperator03.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator03.ColumnCaptionHeight = 17
        Me.CboOperator03.ColumnFooterHeight = 17
        Me.CboOperator03.ContentHeight = 17
        Me.CboOperator03.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator03.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator03.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator03.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator03.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator03.EditorHeight = 17
        Me.CboOperator03.EvenRowStyle = Style58
        Me.CboOperator03.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator03.FooterStyle = Style59
        Me.CboOperator03.HeadingStyle = Style60
        Me.CboOperator03.HighLightRowStyle = Style61
        Me.CboOperator03.Images.Add(CType(resources.GetObject("CboOperator03.Images"), System.Drawing.Image))
        Me.CboOperator03.ItemHeight = 15
        Me.CboOperator03.Location = New System.Drawing.Point(164, -43)
        Me.CboOperator03.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator03.MaxDropDownItems = CType(5, Short)
        Me.CboOperator03.MaxLength = 32767
        Me.CboOperator03.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator03.Name = "CboOperator03"
        Me.CboOperator03.OddRowStyle = Style62
        Me.CboOperator03.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator03.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator03.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator03.SelectedStyle = Style63
        Me.CboOperator03.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator03.Style = Style64
        Me.CboOperator03.TabIndex = 20
        Me.CboOperator03.PropBag = resources.GetString("CboOperator03.PropBag")
        '
        'CboOperator02
        '
        Me.CboOperator02.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator02.Caption = ""
        Me.CboOperator02.CaptionHeight = 17
        Me.CboOperator02.CaptionStyle = Style65
        Me.CboOperator02.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator02.ColumnCaptionHeight = 17
        Me.CboOperator02.ColumnFooterHeight = 17
        Me.CboOperator02.ContentHeight = 17
        Me.CboOperator02.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator02.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator02.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator02.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator02.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator02.EditorHeight = 17
        Me.CboOperator02.EvenRowStyle = Style66
        Me.CboOperator02.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator02.FooterStyle = Style67
        Me.CboOperator02.HeadingStyle = Style68
        Me.CboOperator02.HighLightRowStyle = Style69
        Me.CboOperator02.Images.Add(CType(resources.GetObject("CboOperator02.Images"), System.Drawing.Image))
        Me.CboOperator02.ItemHeight = 15
        Me.CboOperator02.Location = New System.Drawing.Point(162, -45)
        Me.CboOperator02.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator02.MaxDropDownItems = CType(5, Short)
        Me.CboOperator02.MaxLength = 32767
        Me.CboOperator02.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator02.Name = "CboOperator02"
        Me.CboOperator02.OddRowStyle = Style70
        Me.CboOperator02.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator02.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator02.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator02.SelectedStyle = Style71
        Me.CboOperator02.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator02.Style = Style72
        Me.CboOperator02.TabIndex = 19
        Me.CboOperator02.PropBag = resources.GetString("CboOperator02.PropBag")
        '
        'CboField10
        '
        Me.CboField10.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField10.AutoCompletion = True
        Me.CboField10.AutoDropDown = True
        Me.CboField10.Caption = ""
        Me.CboField10.CaptionHeight = 17
        Me.CboField10.CaptionStyle = Style73
        Me.CboField10.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField10.ColumnCaptionHeight = 17
        Me.CboField10.ColumnFooterHeight = 17
        Me.CboField10.ContentHeight = 17
        Me.CboField10.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField10.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField10.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField10.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField10.EditorHeight = 17
        Me.CboField10.EvenRowStyle = Style74
        Me.CboField10.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField10.FooterStyle = Style75
        Me.CboField10.HeadingStyle = Style76
        Me.CboField10.HighLightRowStyle = Style77
        Me.CboField10.Images.Add(CType(resources.GetObject("CboField10.Images"), System.Drawing.Image))
        Me.CboField10.ItemHeight = 15
        Me.CboField10.Location = New System.Drawing.Point(18, -31)
        Me.CboField10.MatchEntryTimeout = CType(2000, Long)
        Me.CboField10.MaxDropDownItems = CType(5, Short)
        Me.CboField10.MaxLength = 32767
        Me.CboField10.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField10.Name = "CboField10"
        Me.CboField10.OddRowStyle = Style78
        Me.CboField10.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField10.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField10.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField10.SelectedStyle = Style79
        Me.CboField10.Size = New System.Drawing.Size(155, 23)
        Me.CboField10.Style = Style80
        Me.CboField10.TabIndex = 18
        Me.CboField10.PropBag = resources.GetString("CboField10.PropBag")
        '
        'CboField09
        '
        Me.CboField09.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField09.AutoCompletion = True
        Me.CboField09.AutoDropDown = True
        Me.CboField09.Caption = ""
        Me.CboField09.CaptionHeight = 17
        Me.CboField09.CaptionStyle = Style81
        Me.CboField09.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField09.ColumnCaptionHeight = 17
        Me.CboField09.ColumnFooterHeight = 17
        Me.CboField09.ContentHeight = 17
        Me.CboField09.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField09.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField09.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField09.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField09.EditorHeight = 17
        Me.CboField09.EvenRowStyle = Style82
        Me.CboField09.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField09.FooterStyle = Style83
        Me.CboField09.HeadingStyle = Style84
        Me.CboField09.HighLightRowStyle = Style85
        Me.CboField09.Images.Add(CType(resources.GetObject("CboField09.Images"), System.Drawing.Image))
        Me.CboField09.ItemHeight = 15
        Me.CboField09.Location = New System.Drawing.Point(16, -33)
        Me.CboField09.MatchEntryTimeout = CType(2000, Long)
        Me.CboField09.MaxDropDownItems = CType(5, Short)
        Me.CboField09.MaxLength = 32767
        Me.CboField09.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField09.Name = "CboField09"
        Me.CboField09.OddRowStyle = Style86
        Me.CboField09.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField09.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField09.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField09.SelectedStyle = Style87
        Me.CboField09.Size = New System.Drawing.Size(155, 23)
        Me.CboField09.Style = Style88
        Me.CboField09.TabIndex = 17
        Me.CboField09.PropBag = resources.GetString("CboField09.PropBag")
        '
        'CboField08
        '
        Me.CboField08.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField08.AutoCompletion = True
        Me.CboField08.AutoDropDown = True
        Me.CboField08.Caption = ""
        Me.CboField08.CaptionHeight = 17
        Me.CboField08.CaptionStyle = Style89
        Me.CboField08.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField08.ColumnCaptionHeight = 17
        Me.CboField08.ColumnFooterHeight = 17
        Me.CboField08.ContentHeight = 17
        Me.CboField08.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField08.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField08.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField08.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField08.EditorHeight = 17
        Me.CboField08.EvenRowStyle = Style90
        Me.CboField08.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField08.FooterStyle = Style91
        Me.CboField08.HeadingStyle = Style92
        Me.CboField08.HighLightRowStyle = Style93
        Me.CboField08.Images.Add(CType(resources.GetObject("CboField08.Images"), System.Drawing.Image))
        Me.CboField08.ItemHeight = 15
        Me.CboField08.Location = New System.Drawing.Point(14, -35)
        Me.CboField08.MatchEntryTimeout = CType(2000, Long)
        Me.CboField08.MaxDropDownItems = CType(5, Short)
        Me.CboField08.MaxLength = 32767
        Me.CboField08.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField08.Name = "CboField08"
        Me.CboField08.OddRowStyle = Style94
        Me.CboField08.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField08.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField08.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField08.SelectedStyle = Style95
        Me.CboField08.Size = New System.Drawing.Size(155, 23)
        Me.CboField08.Style = Style96
        Me.CboField08.TabIndex = 16
        Me.CboField08.PropBag = resources.GetString("CboField08.PropBag")
        '
        'CboField07
        '
        Me.CboField07.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField07.AutoCompletion = True
        Me.CboField07.AutoDropDown = True
        Me.CboField07.Caption = ""
        Me.CboField07.CaptionHeight = 17
        Me.CboField07.CaptionStyle = Style97
        Me.CboField07.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField07.ColumnCaptionHeight = 17
        Me.CboField07.ColumnFooterHeight = 17
        Me.CboField07.ContentHeight = 17
        Me.CboField07.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField07.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField07.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField07.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField07.EditorHeight = 17
        Me.CboField07.EvenRowStyle = Style98
        Me.CboField07.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField07.FooterStyle = Style99
        Me.CboField07.HeadingStyle = Style100
        Me.CboField07.HighLightRowStyle = Style101
        Me.CboField07.Images.Add(CType(resources.GetObject("CboField07.Images"), System.Drawing.Image))
        Me.CboField07.ItemHeight = 15
        Me.CboField07.Location = New System.Drawing.Point(12, -37)
        Me.CboField07.MatchEntryTimeout = CType(2000, Long)
        Me.CboField07.MaxDropDownItems = CType(5, Short)
        Me.CboField07.MaxLength = 32767
        Me.CboField07.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField07.Name = "CboField07"
        Me.CboField07.OddRowStyle = Style102
        Me.CboField07.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField07.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField07.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField07.SelectedStyle = Style103
        Me.CboField07.Size = New System.Drawing.Size(155, 23)
        Me.CboField07.Style = Style104
        Me.CboField07.TabIndex = 15
        Me.CboField07.PropBag = resources.GetString("CboField07.PropBag")
        '
        'CboField06
        '
        Me.CboField06.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField06.AutoCompletion = True
        Me.CboField06.AutoDropDown = True
        Me.CboField06.Caption = ""
        Me.CboField06.CaptionHeight = 17
        Me.CboField06.CaptionStyle = Style105
        Me.CboField06.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField06.ColumnCaptionHeight = 17
        Me.CboField06.ColumnFooterHeight = 17
        Me.CboField06.ContentHeight = 17
        Me.CboField06.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField06.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField06.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField06.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField06.EditorHeight = 17
        Me.CboField06.EvenRowStyle = Style106
        Me.CboField06.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField06.FooterStyle = Style107
        Me.CboField06.HeadingStyle = Style108
        Me.CboField06.HighLightRowStyle = Style109
        Me.CboField06.Images.Add(CType(resources.GetObject("CboField06.Images"), System.Drawing.Image))
        Me.CboField06.ItemHeight = 15
        Me.CboField06.Location = New System.Drawing.Point(10, -39)
        Me.CboField06.MatchEntryTimeout = CType(2000, Long)
        Me.CboField06.MaxDropDownItems = CType(5, Short)
        Me.CboField06.MaxLength = 32767
        Me.CboField06.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField06.Name = "CboField06"
        Me.CboField06.OddRowStyle = Style110
        Me.CboField06.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField06.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField06.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField06.SelectedStyle = Style111
        Me.CboField06.Size = New System.Drawing.Size(155, 23)
        Me.CboField06.Style = Style112
        Me.CboField06.TabIndex = 14
        Me.CboField06.PropBag = resources.GetString("CboField06.PropBag")
        '
        'CboField05
        '
        Me.CboField05.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField05.AutoCompletion = True
        Me.CboField05.AutoDropDown = True
        Me.CboField05.Caption = ""
        Me.CboField05.CaptionHeight = 17
        Me.CboField05.CaptionStyle = Style113
        Me.CboField05.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField05.ColumnCaptionHeight = 17
        Me.CboField05.ColumnFooterHeight = 17
        Me.CboField05.ContentHeight = 17
        Me.CboField05.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField05.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField05.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField05.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField05.EditorHeight = 17
        Me.CboField05.EvenRowStyle = Style114
        Me.CboField05.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField05.FooterStyle = Style115
        Me.CboField05.HeadingStyle = Style116
        Me.CboField05.HighLightRowStyle = Style117
        Me.CboField05.Images.Add(CType(resources.GetObject("CboField05.Images"), System.Drawing.Image))
        Me.CboField05.ItemHeight = 15
        Me.CboField05.Location = New System.Drawing.Point(8, -41)
        Me.CboField05.MatchEntryTimeout = CType(2000, Long)
        Me.CboField05.MaxDropDownItems = CType(5, Short)
        Me.CboField05.MaxLength = 32767
        Me.CboField05.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField05.Name = "CboField05"
        Me.CboField05.OddRowStyle = Style118
        Me.CboField05.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField05.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField05.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField05.SelectedStyle = Style119
        Me.CboField05.Size = New System.Drawing.Size(155, 23)
        Me.CboField05.Style = Style120
        Me.CboField05.TabIndex = 13
        Me.CboField05.PropBag = resources.GetString("CboField05.PropBag")
        '
        'CboField04
        '
        Me.CboField04.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField04.AutoCompletion = True
        Me.CboField04.AutoDropDown = True
        Me.CboField04.Caption = ""
        Me.CboField04.CaptionHeight = 17
        Me.CboField04.CaptionStyle = Style121
        Me.CboField04.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField04.ColumnCaptionHeight = 17
        Me.CboField04.ColumnFooterHeight = 17
        Me.CboField04.ContentHeight = 17
        Me.CboField04.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField04.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField04.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField04.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField04.EditorHeight = 17
        Me.CboField04.EvenRowStyle = Style122
        Me.CboField04.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField04.FooterStyle = Style123
        Me.CboField04.HeadingStyle = Style124
        Me.CboField04.HighLightRowStyle = Style125
        Me.CboField04.Images.Add(CType(resources.GetObject("CboField04.Images"), System.Drawing.Image))
        Me.CboField04.ItemHeight = 15
        Me.CboField04.Location = New System.Drawing.Point(6, -43)
        Me.CboField04.MatchEntryTimeout = CType(2000, Long)
        Me.CboField04.MaxDropDownItems = CType(5, Short)
        Me.CboField04.MaxLength = 32767
        Me.CboField04.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField04.Name = "CboField04"
        Me.CboField04.OddRowStyle = Style126
        Me.CboField04.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField04.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField04.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField04.SelectedStyle = Style127
        Me.CboField04.Size = New System.Drawing.Size(155, 23)
        Me.CboField04.Style = Style128
        Me.CboField04.TabIndex = 12
        Me.CboField04.PropBag = resources.GetString("CboField04.PropBag")
        '
        'CboField03
        '
        Me.CboField03.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField03.AutoCompletion = True
        Me.CboField03.AutoDropDown = True
        Me.CboField03.Caption = ""
        Me.CboField03.CaptionHeight = 17
        Me.CboField03.CaptionStyle = Style129
        Me.CboField03.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField03.ColumnCaptionHeight = 17
        Me.CboField03.ColumnFooterHeight = 17
        Me.CboField03.ContentHeight = 17
        Me.CboField03.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField03.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField03.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField03.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField03.EditorHeight = 17
        Me.CboField03.EvenRowStyle = Style130
        Me.CboField03.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField03.FooterStyle = Style131
        Me.CboField03.HeadingStyle = Style132
        Me.CboField03.HighLightRowStyle = Style133
        Me.CboField03.Images.Add(CType(resources.GetObject("CboField03.Images"), System.Drawing.Image))
        Me.CboField03.ItemHeight = 15
        Me.CboField03.Location = New System.Drawing.Point(4, -45)
        Me.CboField03.MatchEntryTimeout = CType(2000, Long)
        Me.CboField03.MaxDropDownItems = CType(5, Short)
        Me.CboField03.MaxLength = 32767
        Me.CboField03.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField03.Name = "CboField03"
        Me.CboField03.OddRowStyle = Style134
        Me.CboField03.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField03.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField03.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField03.SelectedStyle = Style135
        Me.CboField03.Size = New System.Drawing.Size(155, 23)
        Me.CboField03.Style = Style136
        Me.CboField03.TabIndex = 11
        Me.CboField03.PropBag = resources.GetString("CboField03.PropBag")
        '
        'CboField02
        '
        Me.CboField02.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField02.AutoCompletion = True
        Me.CboField02.AutoDropDown = True
        Me.CboField02.Caption = ""
        Me.CboField02.CaptionHeight = 17
        Me.CboField02.CaptionStyle = Style137
        Me.CboField02.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField02.ColumnCaptionHeight = 17
        Me.CboField02.ColumnFooterHeight = 17
        Me.CboField02.ContentHeight = 17
        Me.CboField02.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField02.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField02.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField02.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField02.EditorHeight = 17
        Me.CboField02.EvenRowStyle = Style138
        Me.CboField02.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField02.FooterStyle = Style139
        Me.CboField02.HeadingStyle = Style140
        Me.CboField02.HighLightRowStyle = Style141
        Me.CboField02.Images.Add(CType(resources.GetObject("CboField02.Images"), System.Drawing.Image))
        Me.CboField02.ItemHeight = 15
        Me.CboField02.Location = New System.Drawing.Point(2, -47)
        Me.CboField02.MatchEntryTimeout = CType(2000, Long)
        Me.CboField02.MaxDropDownItems = CType(5, Short)
        Me.CboField02.MaxLength = 32767
        Me.CboField02.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField02.Name = "CboField02"
        Me.CboField02.OddRowStyle = Style142
        Me.CboField02.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField02.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField02.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField02.SelectedStyle = Style143
        Me.CboField02.Size = New System.Drawing.Size(155, 23)
        Me.CboField02.Style = Style144
        Me.CboField02.TabIndex = 10
        Me.CboField02.PropBag = resources.GetString("CboField02.PropBag")
        '
        'CboOperator01
        '
        Me.CboOperator01.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboOperator01.Caption = ""
        Me.CboOperator01.CaptionHeight = 17
        Me.CboOperator01.CaptionStyle = Style145
        Me.CboOperator01.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboOperator01.ColumnCaptionHeight = 17
        Me.CboOperator01.ColumnFooterHeight = 17
        Me.CboOperator01.ContentHeight = 17
        Me.CboOperator01.DataMode = C1.Win.C1List.DataModeEnum.AddItem
        Me.CboOperator01.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboOperator01.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboOperator01.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator01.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboOperator01.EditorHeight = 17
        Me.CboOperator01.EvenRowStyle = Style146
        Me.CboOperator01.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboOperator01.FooterStyle = Style147
        Me.CboOperator01.HeadingStyle = Style148
        Me.CboOperator01.HighLightRowStyle = Style149
        Me.CboOperator01.Images.Add(CType(resources.GetObject("CboOperator01.Images"), System.Drawing.Image))
        Me.CboOperator01.ItemHeight = 15
        Me.CboOperator01.Location = New System.Drawing.Point(171, 28)
        Me.CboOperator01.MatchEntryTimeout = CType(2000, Long)
        Me.CboOperator01.MaxDropDownItems = CType(5, Short)
        Me.CboOperator01.MaxLength = 32767
        Me.CboOperator01.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboOperator01.Name = "CboOperator01"
        Me.CboOperator01.OddRowStyle = Style150
        Me.CboOperator01.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboOperator01.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboOperator01.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboOperator01.SelectedStyle = Style151
        Me.CboOperator01.Size = New System.Drawing.Size(132, 23)
        Me.CboOperator01.Style = Style152
        Me.CboOperator01.TabIndex = 8
        Me.CboOperator01.PropBag = resources.GetString("CboOperator01.PropBag")
        '
        'CboField01
        '
        Me.CboField01.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.CboField01.AutoCompletion = True
        Me.CboField01.AutoDropDown = True
        Me.CboField01.Caption = ""
        Me.CboField01.CaptionHeight = 17
        Me.CboField01.CaptionStyle = Style153
        Me.CboField01.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CboField01.ColumnCaptionHeight = 17
        Me.CboField01.ColumnFooterHeight = 17
        Me.CboField01.ColumnWidth = 100
        Me.CboField01.ContentHeight = 17
        Me.CboField01.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.CboField01.DropDownWidth = 300
        Me.CboField01.EditorBackColor = System.Drawing.SystemColors.Window
        Me.CboField01.EditorFont = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField01.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.CboField01.EditorHeight = 17
        Me.CboField01.EvenRowStyle = Style154
        Me.CboField01.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboField01.FooterStyle = Style155
        Me.CboField01.HeadingStyle = Style156
        Me.CboField01.HighLightRowStyle = Style157
        Me.CboField01.Images.Add(CType(resources.GetObject("CboField01.Images"), System.Drawing.Image))
        Me.CboField01.ItemHeight = 15
        Me.CboField01.LimitToList = True
        Me.CboField01.Location = New System.Drawing.Point(12, 28)
        Me.CboField01.MatchEntryTimeout = CType(2000, Long)
        Me.CboField01.MaxDropDownItems = CType(5, Short)
        Me.CboField01.MaxLength = 32767
        Me.CboField01.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.CboField01.Name = "CboField01"
        Me.CboField01.OddRowStyle = Style158
        Me.CboField01.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.CboField01.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.CboField01.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.CboField01.SelectedStyle = Style159
        Me.CboField01.Size = New System.Drawing.Size(155, 23)
        Me.CboField01.Style = Style160
        Me.CboField01.TabIndex = 7
        Me.CboField01.PropBag = resources.GetString("CboField01.PropBag")
        '
        'TxtValueB01
        '
        Me.TxtValueB01.Font = New System.Drawing.Font("Lemon3", 8.249999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValueB01.Location = New System.Drawing.Point(303, 29)
        Me.TxtValueB01.Name = "TxtValueB01"
        Me.TxtValueB01.Size = New System.Drawing.Size(211, 22)
        Me.TxtValueB01.TabIndex = 6
        '
        'c1dateA01
        '
        Me.c1dateA01.AutoSize = False
        Me.c1dateA01.CustomFormat = "dd/MM/yyyy"
        Me.c1dateA01.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateA01.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateA01.Location = New System.Drawing.Point(303, 29)
        Me.c1dateA01.Name = "c1dateA01"
        Me.c1dateA01.Size = New System.Drawing.Size(103, 22)
        Me.c1dateA01.TabIndex = 75
        Me.c1dateA01.Tag = Nothing
        Me.c1dateA01.TrimStart = True
        Me.c1dateA01.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'c1dateB01
        '
        Me.c1dateB01.AutoSize = False
        Me.c1dateB01.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB01.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB01.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB01.Location = New System.Drawing.Point(411, 29)
        Me.c1dateB01.Name = "c1dateB01"
        Me.c1dateB01.Size = New System.Drawing.Size(103, 22)
        Me.c1dateB01.TabIndex = 76
        Me.c1dateB01.Tag = Nothing
        Me.c1dateB01.TrimStart = True
        Me.c1dateB01.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'ButAnd
        '
        Me.ButAnd.BackColor = System.Drawing.SystemColors.Control
        Me.ButAnd.Location = New System.Drawing.Point(125, 87)
        Me.ButAnd.Name = "ButAnd"
        Me.ButAnd.Size = New System.Drawing.Size(56, 24)
        Me.ButAnd.TabIndex = 2
        Me.ButAnd.Text = "&Và (F9)"
        Me.ButAnd.UseVisualStyleBackColor = False
        '
        'ButClose
        '
        Me.ButClose.BackColor = System.Drawing.SystemColors.Control
        Me.ButClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButClose.Location = New System.Drawing.Point(463, 87)
        Me.ButClose.Name = "ButClose"
        Me.ButClose.Size = New System.Drawing.Size(70, 24)
        Me.ButClose.TabIndex = 7
        Me.ButClose.Text = "Đó&ng"
        Me.ButClose.UseVisualStyleBackColor = False
        '
        'ButFind
        '
        Me.ButFind.BackColor = System.Drawing.SystemColors.Control
        Me.ButFind.Location = New System.Drawing.Point(392, 87)
        Me.ButFind.Name = "ButFind"
        Me.ButFind.Size = New System.Drawing.Size(70, 24)
        Me.ButFind.TabIndex = 6
        Me.ButFind.Text = "Tìm &Kiếm"
        Me.ButFind.UseVisualStyleBackColor = False
        '
        'ButOr
        '
        Me.ButOr.BackColor = System.Drawing.SystemColors.Control
        Me.ButOr.Location = New System.Drawing.Point(182, 87)
        Me.ButOr.Name = "ButOr"
        Me.ButOr.Size = New System.Drawing.Size(70, 24)
        Me.ButOr.TabIndex = 3
        Me.ButOr.Text = "&Hoặc (F10)"
        Me.ButOr.UseVisualStyleBackColor = False
        '
        'c1dateB09
        '
        Me.c1dateB09.AutoSize = False
        Me.c1dateB09.CustomFormat = "dd/MM/yyyy"
        Me.c1dateB09.Font = New System.Drawing.Font("Lemon3", 8.25!)
        Me.c1dateB09.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateB09.Location = New System.Drawing.Point(324, -34)
        Me.c1dateB09.Name = "c1dateB09"
        Me.c1dateB09.Size = New System.Drawing.Size(100, 22)
        Me.c1dateB09.TabIndex = 9
        Me.c1dateB09.Tag = Nothing
        Me.c1dateB09.TrimStart = True
        Me.c1dateB09.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'btnPrevious
        '
        Me.btnPrevious.BackColor = System.Drawing.SystemColors.Control
        Me.btnPrevious.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPrevious.Location = New System.Drawing.Point(6, 87)
        Me.btnPrevious.Name = "btnPrevious"
        Me.btnPrevious.Size = New System.Drawing.Size(68, 24)
        Me.btnPrevious.TabIndex = 1
        Me.btnPrevious.Text = "&Trở lại (F8)"
        Me.btnPrevious.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrevious.UseVisualStyleBackColor = False
        '
        'btnBackFirst
        '
        Me.btnBackFirst.BackColor = System.Drawing.SystemColors.Control
        Me.btnBackFirst.Image = CType(resources.GetObject("btnBackFirst.Image"), System.Drawing.Image)
        Me.btnBackFirst.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBackFirst.Location = New System.Drawing.Point(304, 87)
        Me.btnBackFirst.Name = "btnBackFirst"
        Me.btnBackFirst.Size = New System.Drawing.Size(50, 24)
        Me.btnBackFirst.TabIndex = 5
        Me.btnBackFirst.Text = "   (F12)"
        Me.btnBackFirst.UseVisualStyleBackColor = False
        '
        'ButBack
        '
        Me.ButBack.BackColor = System.Drawing.SystemColors.Control
        Me.ButBack.Image = CType(resources.GetObject("ButBack.Image"), System.Drawing.Image)
        Me.ButBack.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButBack.Location = New System.Drawing.Point(253, 87)
        Me.ButBack.Name = "ButBack"
        Me.ButBack.Size = New System.Drawing.Size(50, 24)
        Me.ButBack.TabIndex = 4
        Me.ButBack.Text = "   (F11)"
        Me.ButBack.UseVisualStyleBackColor = False
        '
        'D99F0006
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(540, 118)
        Me.Controls.Add(Me.btnPrevious)
        Me.Controls.Add(Me.btnBackFirst)
        Me.Controls.Add(Me.c1dateB09)
        Me.Controls.Add(Me.ButBack)
        Me.Controls.Add(Me.ButOr)
        Me.Controls.Add(Me.ButFind)
        Me.Controls.Add(Me.ButClose)
        Me.Controls.Add(Me.ButAnd)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D99F0006"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "TØm kiÕm"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.c1dateB10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB08, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB07, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB06, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB05, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB04, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB03, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB02, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA09, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA08, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA07, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA06, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA05, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA04, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA03, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA02, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator09, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator08, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator07, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator06, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator05, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator04, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator03, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator02, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField09, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField08, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField07, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField06, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField05, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField04, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField03, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField02, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboOperator01, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboField01, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateA01, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB01, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.c1dateB09, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Const LM_H As Integer = 50
    Public Const LM_TOTAL_ROW As Integer = 10
    Public Const LM_DOUBLE_WIDTH As Integer = 206
    Public Const LM_WIDTH As Integer = 103
    Public Const LM_SPACE As Integer = 5

    'Chiều cao ban đầu của Form trước khi thêm các điều kiện tìm kiếm
    Public Const LM_Form_Height As Integer = 143
    'Chiều cao ban đầu của GroupBox
    Public Const LM_GroupBox_Height As Integer = 70
    'Vị trí ban đầu của các Button
    Public Const LM_Button_Top As Integer = 87

    Private idx As Integer
    Private xa() As XArray
    Public dropH As Integer
    Public LM_AND As String
    Public LM_OR As String
    Public LM_OP_STARTWITH As String
    Public LM_OP_CONTAIN As String
    Public LM_OP_NOTCONTAIN As String
    Public LM_OP_ENDWITH As String
    Public LM_OP_EQUAL As String
    Public LM_OP_GREATEREQUAL As String
    Public LM_OP_GREATER As String
    Public LM_OP_LESSEQUAL As String
    Public LM_OP_LESS As String
    Public LM_OP_BETWEEN As String
    Public LM_OP_NOTBETWEEN As String
    Public LM_OP_NOTEQUAL As String
    Public LM_ASC As String
    Public LM_DESCAS As String
    Public LM_OKCommandCaption As String
    Public LM_DELETECommandCaption As String
    Public LM_CloseCommandCaption As String
    Public LM_FIELDNAMECaption As String
    Public LM_ORDERNOCaption As String

    'Public LM_FindClient As Boolean ' Cờ tìm kiếm client
    Public LM_FindClient As Int16 = 0 '= 0: Tìm kiếm server; = 1: Tìm kiếm client; = 2: Tìm kiếm client nhưng có trả về giá trị server;

    Private bFlagBETWEEN As Boolean

    Dim DataTypeField01 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField02 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField03 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField04 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField05 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField06 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField07 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField08 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField09 As FinderTypeEnum = FinderTypeEnum.lmFinderString
    Dim DataTypeField10 As FinderTypeEnum = FinderTypeEnum.lmFinderString

    Dim bChangeText As Boolean = False
    Dim sMsgNhapSoChuaDung As String
    Dim sMsgNhapSoQuaLon As String

    'TemplateID hiện tại (sẽ bị thay đổi khi bấm nút Trở Lại) 
    Dim iTemplateID As Integer
    'TemplateID của lần tìm kiếm gần nhất (ko bị thay đổi khi bấm nút trở lại)
    Dim iLastTemplateID As Integer
    'TemplateID lớn nhất 
    Dim iMaxTemplateID As Integer

    Private Sub EnableButton(ByVal blnStatus As Boolean)
        'Update 02/07/2010: idx>=10 thì tắt nút And Or
        If idx >= LM_TOTAL_ROW Then
            ButFind.Enabled = False
            ButAnd.Enabled = False
            ButOr.Enabled = False
            'ButBack.Enabled = blnStatus
        Else
            ButFind.Enabled = blnStatus
            ButAnd.Enabled = blnStatus
            ButOr.Enabled = blnStatus
            'ButBack.Enabled = blnStatus
        End If
    End Sub

    Private Sub EnableVisibleControls(ByVal cboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal chk As System.Windows.Forms.CheckBox, ByVal blnStatus As Boolean)
        Try
            If chk.Visible Then chk.Enabled = True

            If TxtValueA.Visible Then
                TxtValueA.Enabled = blnStatus
                If blnStatus Then
                    If bChangeText Then TxtValueA.Text = "" : TxtValueB.Text = ""
                    TxtValueA.Focus()
                End If
                If TxtValueB.Visible Then TxtValueB.Enabled = blnStatus
            Else
                If c1dateA.Visible Then
                    c1dateA.Enabled = blnStatus
                    If blnStatus Then
                        If bChangeText Then c1dateA.Value = Now : c1dateB.Value = Now
                        c1dateA.Focus()
                    End If
                    If c1dateB.Visible Then c1dateB.Enabled = blnStatus
                End If
            End If
            'Hoàng Long: Khi comboOperator được chọn dữ liệu mới sáng butFind
            If CboOperator.Text <> "" Then
                ButFind.Enabled = True
            End If
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try

    End Sub

    Private Sub CloseCboField(ByVal CboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal chk As System.Windows.Forms.CheckBox)
        If Trim(CboField.Text) <> "" Then
            bChangeText = False
            EnableButton(True)
            StuffOperatorCombo(CboField, CType(CboField.Columns(2).Value, D99D0041.FinderTypeEnum))

            CloseCboOperator(CboField, CboOperator, TxtValueA, TxtValueB, c1dateA, c1dateB, chk)
        End If
    End Sub

    Private Sub CboField01_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField01.Close
        CloseCboField(CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, chk01)
    End Sub

    Private Sub CboField02_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField02.Close
        CloseCboField(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, chk02)
    End Sub

    Private Sub CboField03_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField03.Close
        CloseCboField(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, chk03)
    End Sub

    Private Sub CboField04_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField04.Close
        CloseCboField(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, chk04)

    End Sub

    Private Sub CboField05_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField05.Close
        CloseCboField(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, chk05)

    End Sub

    Private Sub CboField06_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField06.Close
        CloseCboField(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, chk06)

    End Sub

    Private Sub CboField07_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField07.Close
        CloseCboField(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, chk07)

    End Sub

    Private Sub CboField08_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField08.Close
        CloseCboField(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, chk08)

    End Sub

    Private Sub CboField09_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField09.Close
        CloseCboField(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, chk09)

    End Sub

    Private Sub CboField10_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboField10.Close
        CloseCboField(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, chk10)

    End Sub

    Private Sub CboField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField01.KeyDown, CboField02.KeyDown, CboField03.KeyDown, CboField04.KeyDown, CboField05.KeyDown, CboField06.KeyDown, CboField07.KeyDown, CboField08.KeyDown, CboField09.KeyDown, CboField10.KeyDown
        Dim tdbc As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        Select Case e.KeyCode
            Case Keys.A, Keys.D, Keys.E, Keys.I, Keys.O, Keys.U, Keys.Y, Keys.Back
                tdbc.AutoCompletion = False
            Case Else
                tdbc.AutoCompletion = True
        End Select
    End Sub


    'Private Sub CboField01_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField01.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01)
    '    End If
    'End Sub

    'Private Sub CboField02_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField02.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02)
    '    End If
    'End Sub

    'Private Sub CboField03_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField03.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03)
    '    End If
    'End Sub

    'Private Sub CboField04_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField04.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04)
    '    End If
    'End Sub

    'Private Sub CboField05_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField05.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05)
    '    End If
    'End Sub

    'Private Sub CboField06_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField06.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06)
    '    End If
    'End Sub

    'Private Sub CboField07_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField07.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07)
    '    End If
    'End Sub

    'Private Sub CboField08_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField08.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08)
    '    End If
    'End Sub

    'Private Sub CboField09_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField09.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09)
    '    End If
    'End Sub

    'Private Sub CboField10_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboField10.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        EnabledControl_CboFieldKeyDelete(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10)
    '    End If
    'End Sub

    Private Sub CboField01_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField01.LostFocus
        'Update 07/07/2010
        If CboField01.Columns("Col1").Text.Contains(CboField01.Text) Then
            CboField01.AutoCompletion = True
            CboField01.Text = CboField01.Columns("Col1").Text
        End If

        If CboField01.FindStringExact(CboField01.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, chk01)
        End If
    End Sub

    Private Sub CboField02_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField02.LostFocus
        'Update 07/07/2010
        If CboField02.Columns("Col1").Text.Contains(CboField02.Text) Then
            CboField02.AutoCompletion = True
            CboField02.Text = CboField02.Columns("Col1").Text
        End If

        If CboField02.FindStringExact(CboField02.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, chk02)
        End If
    End Sub

    Private Sub CboField03_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField03.LostFocus
        'Update 07/07/2010
        If CboField03.Columns("Col1").Text.Contains(CboField03.Text) Then
            CboField03.AutoCompletion = True
            CboField03.Text = CboField03.Columns("Col1").Text
        End If

        If CboField03.FindStringExact(CboField03.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, chk03)
        End If
    End Sub

    Private Sub CboField04_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField04.LostFocus
        'Update 07/07/2010
        If CboField04.Columns("Col1").Text.Contains(CboField04.Text) Then
            CboField04.AutoCompletion = True
            CboField04.Text = CboField04.Columns("Col1").Text
        End If
        If CboField04.FindStringExact(CboField04.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, chk04)
        End If
    End Sub

    Private Sub CboField05_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField05.LostFocus
        'Update 07/07/2010
        If CboField05.Columns("Col1").Text.Contains(CboField05.Text) Then
            CboField05.AutoCompletion = True
            CboField05.Text = CboField05.Columns("Col1").Text
        End If

        If CboField05.FindStringExact(CboField05.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, chk05)
        End If
    End Sub

    Private Sub CboField06_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField06.LostFocus
        'Update 07/07/2010
        If CboField06.Columns("Col1").Text.Contains(CboField06.Text) Then
            CboField06.AutoCompletion = True
            CboField06.Text = CboField06.Columns("Col1").Text
        End If

        If CboField06.FindStringExact(CboField06.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, chk06)
        End If
    End Sub

    Private Sub CboField07_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField07.LostFocus
        'Update 07/07/2010
        If CboField07.Columns("Col1").Text.Contains(CboField07.Text) Then
            CboField07.AutoCompletion = True
            CboField07.Text = CboField07.Columns("Col1").Text
        End If

        If CboField07.FindStringExact(CboField07.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, chk07)
        End If
    End Sub

    Private Sub CboField08_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField08.LostFocus
        'Update 07/07/2010
        If CboField08.Columns("Col1").Text.Contains(CboField08.Text) Then
            CboField08.AutoCompletion = True
            CboField08.Text = CboField08.Columns("Col1").Text
        End If

        If CboField08.FindStringExact(CboField08.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, chk08)
        End If
    End Sub

    Private Sub CboField09_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField09.LostFocus
        'Update 07/07/2010
        If CboField09.Columns("Col1").Text.Contains(CboField09.Text) Then
            CboField09.AutoCompletion = True
            CboField09.Text = CboField09.Columns("Col1").Text
        End If

        If CboField09.FindStringExact(CboField09.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, chk09)
        End If
    End Sub

    Private Sub CboField10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboField10.LostFocus
        'Update 07/07/2010
        If CboField10.Columns("Col1").Text.Contains(CboField10.Text) Then
            CboField10.AutoCompletion = True
            CboField10.Text = CboField10.Columns("Col1").Text
        End If

        If CboField10.FindStringExact(CboField10.Text) = -1 Then
            EnabledControl_CboFieldKeyDelete(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, chk10)
        End If
    End Sub

    Private Sub CloseCboOperator(ByVal cboField As C1.Win.C1List.C1Combo, ByVal cboOperator As C1.Win.C1List.C1Combo, ByVal txtValueA As System.Windows.Forms.TextBox, ByVal txtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal chk As System.Windows.Forms.CheckBox)
        Try
            bFlagBETWEEN = False

            txtValueA.Visible = False
            txtValueB.Visible = False
            c1dateA.Visible = False
            c1dateB.Visible = False
            chk.Visible = False

            Select Case cboOperator.Text

                Case LM_OP_BETWEEN, LM_OP_NOTBETWEEN

                    Select Case CType(cboField.Columns(2).Value, D99D0041.FinderTypeEnum)
                        Case FinderTypeEnum.lmFinderDate
                            c1dateA.Width = LM_WIDTH
                            c1dateA.Left = cboOperator.Left + cboOperator.Width + LM_SPACE
                            c1dateA.Visible = True

                            c1dateB.Width = LM_WIDTH
                            c1dateB.Left = c1dateA.Left + c1dateA.Width + LM_SPACE
                            c1dateB.Visible = True

                        Case Else
                            txtValueA.Width = LM_WIDTH
                            txtValueA.Left = cboOperator.Left + cboOperator.Width + LM_SPACE
                            txtValueA.TextAlign = HorizontalAlignment.Right
                            txtValueA.MaxLength = 0
                            txtValueA.Visible = True

                            txtValueB.Width = LM_WIDTH
                            txtValueB.Left = txtValueA.Left + txtValueA.Width + LM_SPACE
                            txtValueB.TextAlign = HorizontalAlignment.Right
                            txtValueB.Visible = True

                    End Select
                    bFlagBETWEEN = True

                Case Else
                    Select Case CType(cboField.Columns(2).Value, D99D0041.FinderTypeEnum)

                        Case FinderTypeEnum.lmFinderDate
                            c1dateA.Width = LM_DOUBLE_WIDTH + LM_SPACE
                            c1dateA.Left = cboOperator.Left + cboOperator.Width + LM_SPACE
                            c1dateA.Visible = True

                        Case FinderTypeEnum.lmFinderString
                            txtValueA.Width = LM_DOUBLE_WIDTH + LM_SPACE
                            txtValueA.Left = cboOperator.Left + cboOperator.Width + LM_SPACE
                            txtValueA.TextAlign = HorizontalAlignment.Left
                            'Update 16/05/2011: Không giới hạn dữ liệu nhập vào, để người dùng có thể tìm theo tập hợp: Nguyen Van A;Nguyen Thi B
                            'txtValueA.MaxLength = CInt(cboField.Columns(3).Value)
                            txtValueA.Visible = True

                        Case FinderTypeEnum.lmFinderTinyInt
                            'Update 02/07/2010: Số là TinyInt (là checkbox trên lưới)
                            txtValueA.Visible = False
                            chk.Visible = True
                            chk.Left = cboOperator.Left + cboOperator.Width + LM_SPACE

                        Case Else ' Số
                            txtValueA.Width = LM_DOUBLE_WIDTH + LM_SPACE
                            txtValueA.Left = cboOperator.Left + cboOperator.Width + LM_SPACE
                            txtValueA.TextAlign = HorizontalAlignment.Right
                            txtValueA.MaxLength = 0
                            txtValueA.Visible = True
                    End Select

            End Select
            EnableVisibleControls(cboField, cboOperator, txtValueA, txtValueB, c1dateA, c1dateB, chk, True)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try


    End Sub

    Private Sub CboOperator1_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator01.Close
        bChangeText = False
        CloseCboOperator(CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, chk01)
    End Sub

    Private Sub CboOperator2_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator02.Close
        bChangeText = False
        CloseCboOperator(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, chk02)
    End Sub

    Private Sub CboOperator3_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator03.Close
        bChangeText = False
        CloseCboOperator(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, chk03)
    End Sub

    Private Sub CboOperator4_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator04.Close
        bChangeText = False
        CloseCboOperator(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, chk04)

    End Sub

    Private Sub CboOperator5_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator05.Close
        bChangeText = False
        CloseCboOperator(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, chk05)
    End Sub

    Private Sub CboOperator6_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator06.Close
        bChangeText = False
        CloseCboOperator(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, chk06)

    End Sub

    Private Sub CboOperator7_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator07.Close
        bChangeText = False
        CloseCboOperator(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, chk07)
    End Sub

    Private Sub CboOperator8_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator08.Close
        bChangeText = False
        CloseCboOperator(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, chk08)

    End Sub

    Private Sub CboOperator9_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator09.Close
        CloseCboOperator(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, chk09)

    End Sub

    Private Sub CboOperator10_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboOperator10.Close
        bChangeText = False
        CloseCboOperator(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, chk10)

    End Sub

    Private Sub ButAnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButAnd.Click
        AddRow(LM_AND)
        EnableButton(False)
        Center(Me)
    End Sub

    Private Sub ButClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButClose.Click
        strSQLAdvanced = ""
        Me.Close()
    End Sub

    Private Sub ButBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButBack.Click
        DeleteRow()
        Center(Me)
    End Sub

    Private Sub SplitClause()

    End Sub

    Private Sub ListFind(ByVal cboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal Label As System.Windows.Forms.Label, ByVal chk As System.Windows.Forms.CheckBox)
        Dim sValueA As String
        Dim sValueB As String
        Dim findex As Integer
        Dim sClause As String
        'Dim a As New D99C0002

        sClause = ""
        findex = CType(cboField.Columns(2).Text, Integer)

        Select Case findex

            Case FinderTypeEnum.lmFinderString
                'only one text field
                If TxtValueA.Text <> "" Then
                    'If _useUnicode Then
                    '    sValueA = ConvertUnicodeToVni(TxtValueA.Text.Trim)
                    'Else
                    sValueA = TxtValueA.Text.Trim
                    'End If

                    'remove quote
                    sValueA = sValueA.Replace("'", "''")
                    If CboOperator.Text <> LM_OP_EQUAL Then
                        'Update 29/06/2010: CboOperator.Text = thì không Replace
                        'Tìm ký tự đặc biệt % thay bằng [%]
                        sValueA = sValueA.Replace("%", "[%]")
                        'Tìm ký tự đặc biệt * thay bằng [*]
                        sValueA = sValueA.Replace("*", "[*]")
                    End If

                    'Minh Hòa update 09/06/2009: Tìm kiếm theo tập hợp
                    Dim sChar As String = ";"
                    Dim sArr() As String = {""}
                    Dim sClauseT As String = ""
                    Dim bSplit As Boolean = False
                    If sValueA.Contains(sChar) Then
                        'Tìm theo tập hợp
                        sArr = Microsoft.VisualBasic.Split(sValueA, sChar)
                        bSplit = True
                    End If

                    Select Case CboOperator.Text
                        'Mặc định là Có chứa
                        Case LM_OP_CONTAIN
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= cboField.Columns(1).Text & " Like N'%" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " Or " & cboField.Columns(1).Text & " Like N'%" & sArr(i) & "%'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "( " & cboField.Columns(1).Text & " Like N'%" & sValueA & "%')"
                            End If
                            'Update 12/07/2010: incident 33798
                        Case LM_OP_NOTCONTAIN
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= cboField.Columns(1).Text & " Not Like N'%" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " And " & cboField.Columns(1).Text & " Not Like N'%" & sArr(i) & "%'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "( " & cboField.Columns(1).Text & " Not Like N'%" & sValueA & "%')"
                            End If
                        Case LM_OP_STARTWITH
                            If bSplit Then
                                'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= cboField.Columns(1).Text & " Like N'" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " Or " & cboField.Columns(1).Text & " Like N'" & sArr(i) & "%'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "( " & cboField.Columns(1).Text & " Like N'" & sValueA & "%')"
                            End If
                        Case LM_OP_ENDWITH
                            If bSplit Then
                                'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= cboField.Columns(1).Text & " Like N'%" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & cboField.Columns(1).Text & " Like N'%" & sArr(i) & "'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "(" & cboField.Columns(1).Text & " Like N'%" & sValueA & "')"
                            End If

                        Case LM_OP_EQUAL
                            If bSplit Then
                                'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= cboField.Columns(1).Text & " = N'" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & cboField.Columns(1).Text & " = N'" & sArr(i) & "'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "(" & cboField.Columns(1).Text & " = N'" & sValueA & "')"
                            End If

                        Case LM_OP_NOTEQUAL
                            If bSplit Then
                                'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= cboField.Columns(1).Text & " Not Like N'" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & cboField.Columns(1).Text & " Not Like N'" & sArr(i) & "'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "(" & cboField.Columns(1).Text & " Not Like N'" & sValueA & "')"
                            End If
                    End Select
                End If

            Case FinderTypeEnum.lmFinderDate 'Thay đổi
                'Date: 11/03/2008 Dùng hàm SQLDateSave
                'Mặc định là Trong khoảng
                Select Case CboOperator.Text
                    Case LM_OP_BETWEEN
                        c1dateB.UpdateValueWithCurrentText()
                        sClause = "(" & cboField.Columns(1).Text & " >= " & SQLDateSave(c1dateA.Value)
                        sClause = sClause & " And " & cboField.Columns(1).Text & " - 1 < " & SQLDateSave(c1dateB.Value) & ")"
                    Case LM_OP_EQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & cboField.Columns(1).Text & " >= " & SQLDateSave(c1dateA.Value)
                        sClause = sClause & " And " & cboField.Columns(1).Text & " - 1 < " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_GREATEREQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & cboField.Columns(1).Text & " >= " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_GREATER
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & cboField.Columns(1).Text & " - 1 >= " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_LESSEQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & cboField.Columns(1).Text & " <= " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_LESS
                        sClause = "(" & cboField.Columns(1).Text & " < " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_NOTBETWEEN
                        c1dateB.UpdateValueWithCurrentText()
                        sClause = "(" & cboField.Columns(1).Text & " < " & SQLDateSave(c1dateA.Value)
                        sClause = sClause & " Or " & cboField.Columns(1).Text & " - 1 >= " & SQLDateSave(c1dateB.Value) & ")"
                    Case LM_OP_NOTEQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & cboField.Columns(1).Text & " < " & SQLDateSave(c1dateA.Value)
                        sClause = sClause & " Or " & cboField.Columns(1).Text & " - 1 >= " & SQLDateSave(c1dateA.Value) & ")"

                End Select

            Case FinderTypeEnum.lmFinderTinyInt ' Update 02/07/2010: checkbox
                sClause = "( " & cboField.Columns(1).Text & " = " & SQLNumber(chk.Checked) & ")"


            Case FinderTypeEnum.lmFinderInt, FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                '=============================================
                '=='Bổ sung tìm kiếm % - update 07/11/2012====
                '=============================================
                sValueA = TxtValueA.Text
                sValueB = TxtValueB.Text
                If cboField.Columns(4).Text = "Percent" Then
                    If IsNumeric(sValueA) Then sValueA = (Number(sValueA) / 100).ToString
                    If IsNumeric(sValueB) Then sValueB = (Number(sValueB) / 100).ToString
                End If
                '===============================================
                Select Case CboOperator.Text
                    'Mặc định là Bằng
                    Case LM_OP_EQUAL
                        If IsNumeric(sValueA) Then
                            sClause = "( " & cboField.Columns(1).Text & " = " & sValueA & ")"
                        End If
                    Case LM_OP_GREATEREQUAL
                        If IsNumeric(sValueA) Then
                            sClause = "( " & cboField.Columns(1).Text & " >= " & sValueA & ")"
                        End If
                    Case LM_OP_GREATER
                        If IsNumeric(sValueA) Then
                            sClause = "( " & cboField.Columns(1).Text & " > " & sValueA & ")"
                        End If
                    Case LM_OP_LESSEQUAL
                        If IsNumeric(sValueA) Then
                            sClause = "( " & cboField.Columns(1).Text & " <= " & sValueA & ")"
                        End If
                    Case LM_OP_LESS
                        If IsNumeric(sValueA) Then
                            sClause = "( " & cboField.Columns(1).Text & " < " & sValueA & ")"
                        End If
                    Case LM_OP_BETWEEN 'Thay đổi
                        If IsNumeric(sValueA) Then
                            sClause = "( " & cboField.Columns(1).Text & " between " & sValueA
                            sClause = sClause & " and " & sValueB & ")"
                        End If
                    Case LM_OP_NOTBETWEEN 'Thay đổi
                        If IsNumeric(sValueA) And IsNumeric(sValueB) Then
                            sClause = "( " & cboField.Columns(1).Text & " not between " & sValueA
                            sClause = sClause & " and " & sValueB & ")"
                        End If
                    Case LM_OP_NOTEQUAL
                        If IsNumeric(sValueA) Then
                            sClause = "( " & cboField.Columns(1).Text & " <> " & sValueA & ")"
                        End If

                End Select

        End Select


        sClause = sClause.Trim

        If sClause <> "" Then
            If strSQLAdvanced <> "" Then

                If Label.Text = LM_AND Then
                    strSQLAdvanced = strSQLAdvanced & " AND " & sClause
                Else
                    strSQLAdvanced = strSQLAdvanced & " OR " & sClause
                End If
            Else
                strSQLAdvanced = sClause
            End If

        End If

        'a = Nothing
    End Sub

    Private Sub ListFindClient(ByVal cboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal Label As System.Windows.Forms.Label, ByVal chk As System.Windows.Forms.CheckBox)
        Dim sValueA As String
        Dim sValueB As String
        Dim findex As Integer
        Dim sClause As String
        Dim sClauseForReport As String

        sClause = ""
        sClauseForReport = ""
        findex = CType(cboField.Columns(2).Text, Integer)
        Select Case findex
            Case FinderTypeEnum.lmFinderString
                'only one text field             
                If TxtValueA.Text <> "" Then
                    sValueA = TxtValueA.Text.Trim
                    'remove quote
                    sValueA = sValueA.Replace("'", "''")
                    If CboOperator.Text <> LM_OP_EQUAL Then
                        'Update 29/06/2010: CboOperator.Text = thì không Replace
                        'Tìm ký tự đặc biệt % thay bằng [%]
                        sValueA = sValueA.Replace("%", "[%]")
                        'Tìm ký tự đặc biệt * thay bằng [*]
                        sValueA = sValueA.Replace("*", "[*]")
                    End If

                    'Minh Hòa update 09/06/2009: Tìm kiếm theo tập hợp
                    Dim sChar As String = ";"
                    Dim sArr() As String = {""}
                    Dim sClauseT As String = ""
                    Dim bSplit As Boolean = False
                    If sValueA.Contains(sChar) Then
                        'Tìm theo tập hợp
                        sArr = Microsoft.VisualBasic.Split(sValueA, sChar)
                        bSplit = True
                    End If

                    Select Case CboOperator.Text
                        'Mặc định là Có chứa
                        Case LM_OP_CONTAIN
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        'sArr(i) = sArr(i).Replace("*", "[*]")

                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Like '%" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like '%" & sArr(i) & "%'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " Like '%" & sValueA & "%')"
                                sClauseForReport = "( " & "[" & cboField.Columns(1).Text & "]" & " Like N'%" & sValueA & "%')"
                            End If

                            'Update 12/07/2010: incident 33798
                        Case LM_OP_NOTCONTAIN
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        'sArr(i) = sArr(i).Replace("*", "[*]")

                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Not Like '%" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " And " & "[" & cboField.Columns(1).Text & "]" & " Not Like '%" & sArr(i) & "%'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " Not Like '%" & sValueA & "%')"
                            End If

                        Case LM_OP_STARTWITH
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Like '" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like '" & sArr(i) & "%'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " Like '" & sValueA & "%')"
                            End If

                        Case LM_OP_ENDWITH
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Like '%" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like '%" & sArr(i) & "'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " Like '%" & sValueA & "')"
                            End If

                        Case LM_OP_EQUAL
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " = '" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " = '" & sArr(i) & "'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " = '" & sValueA & "')"
                                If sValueA = "" Then sClause = "(" & sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & " is NULL)"
                            End If

                        Case LM_OP_NOTEQUAL
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Not Like '" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Not Like '" & sArr(i) & "'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                            Else
                                sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " Not Like '" & sValueA & "')"
                                If sValueA = "" Then sClause = "(" & sClause & " And " & "[" & cboField.Columns(1).Text & "]" & " is not NULL)"
                            End If

                    End Select
                End If

            Case FinderTypeEnum.lmFinderDate
                ' Date: 05/12/2007 Thay thế các hàm SQLDateSave  thành DateSave
                'Mặc định là Trong khoảng
                Select Case CboOperator.Text
                    Case LM_OP_BETWEEN
                        c1dateB.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >= #" & DateSave(c1dateA.Value) & "#"
                        'sClause = sClause & " And " & cboField.Columns(1).Text & " <= #" & DateSave(c1dateB.Value) & "#)"
                        sClause = sClause & " And " & "[" & cboField.Columns(1).Text & "]" & " < #" & DateSave(CDate(c1dateB.Value).AddDays(1)) & "#)"
                    Case LM_OP_EQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >= #" & DateSave(c1dateA.Value) & "#"
                        sClause = sClause & " And " & "[" & cboField.Columns(1).Text & "]" & "  < #" & DateSave(CDate(c1dateA.Value).AddDays(1)) & "# )"
                    Case LM_OP_GREATEREQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >= #" & DateSave(c1dateA.Value) & "#)"
                    Case LM_OP_GREATER
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >= " & " #" & DateSave(CDate(c1dateA.Value).AddDays(1)) & "#)"
                    Case LM_OP_LESSEQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " <= #" & DateSave(c1dateA.Value) & "#)"
                    Case LM_OP_LESS
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " < #" & DateSave(c1dateA.Value) & "#)"
                    Case LM_OP_NOTBETWEEN
                        c1dateB.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " < #" & DateSave(c1dateA.Value) & "#"
                        sClause = sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & "  >= #" & DateSave(CDate(c1dateB.Value).AddDays(1)) & "#)"
                    Case LM_OP_NOTEQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " < #" & DateSave(c1dateA.Value) & "#"
                        sClause = sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & "  >= #" & DateSave(CDate(c1dateA.Value).AddDays(1)) & "# )"
                End Select

            Case FinderTypeEnum.lmFinderTinyInt 'Update 02/07/2010: Checkbox
                sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " = " & SQLNumber(chk.Checked) & ")"

            Case FinderTypeEnum.lmFinderInt, FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                '=============================================
                '=='Bổ sung tìm kiếm % - update 07/11/2012====
                '=============================================
                sValueA = TxtValueA.Text
                sValueB = TxtValueB.Text
                If cboField.Columns(4).Text = "Percent" Then
                    If IsNumeric(sValueA) Then sValueA = (Number(sValueA) / 100).ToString
                    If IsNumeric(sValueB) Then sValueB = (Number(sValueB) / 100).ToString
                End If
                '===============================================
                Select Case CboOperator.Text
                    'Mặc định là bằng
                    Case LM_OP_EQUAL
                        If IsNumeric(sValueA) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " = " & CDbl(sValueA) & ")"
                        End If
                        If Number(sValueA) = 0 And sValueA <> "" Then sClause = "(" & sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & " is NULL )"
                    Case LM_OP_GREATEREQUAL
                        If IsNumeric(sValueA) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " >= " & CDbl(sValueA) & ")"
                        End If
                    Case LM_OP_GREATER
                        If IsNumeric(sValueA) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " > " & CDbl(sValueA) & ")"
                        End If
                    Case LM_OP_LESSEQUAL
                        If IsNumeric(sValueA) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " <= " & CDbl(sValueA) & ")"
                        End If
                    Case LM_OP_LESS
                        If IsNumeric(sValueA) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " < " & CDbl(sValueA) & ")"
                        End If
                    Case LM_OP_BETWEEN
                        If IsNumeric(sValueA) And IsNumeric(sValueB) Then
                            If sValueB = "" Then sValueB = sValueA
                            sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >=" & CDbl(sValueA)
                            sClause = sClause & " And " & "[" & cboField.Columns(1).Text & "]" & " <= " & CDbl(sValueB) & ")"
                        End If
                    Case LM_OP_NOTBETWEEN
                        If IsNumeric(sValueA) And IsNumeric(sValueB) Then
                            sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " <" & CDbl(sValueA)
                            sClause = sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & " > " & CDbl(sValueB) & ")"
                        End If
                    Case LM_OP_NOTEQUAL
                        If IsNumeric(sValueA) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " <> " & CDbl(sValueA) & ")"
                        End If
                        If Number(sValueA) = 0 And sValueA <> "" Then sClause = "(" & sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & " is not NULL )"
                End Select
        End Select
        sClause = sClause.Trim
        sClauseForReport = sClauseForReport.Trim
        If sClause <> "" Then
            If strSQLAdvanced <> "" Then
                If Label.Text = LM_AND Then
                    strSQLAdvanced = strSQLAdvanced & " AND " & sClause
                Else
                    strSQLAdvanced = strSQLAdvanced & " OR " & sClause
                End If
            Else
                strSQLAdvanced = sClause
                strSQLAdvancedForReport = sClauseForReport
            End If
        End If
    End Sub

    Private Sub ListFindClientServer(ByVal cboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal Label As System.Windows.Forms.Label, ByVal chk As System.Windows.Forms.CheckBox)
        'Update 07/10/2010: Tìm kiếm client nhưng có trả về giá trị cho server
        Dim sValueA As String
        Dim findex As Integer
        Dim sClause As String ' Update 07/10/2010: Trả về chuỗi tìm kiếm cho Client dùng
        Dim sClauseServer As String ' Update 07/10/2010: Trả về chuỗi tìm kiếm cho Server dùng

        sClause = ""
        sClauseServer = ""
        findex = CType(cboField.Columns(2).Text, Integer)
        Select Case findex

            Case FinderTypeEnum.lmFinderString
                'only one text field             
                If TxtValueA.Text <> "" Then
                    sValueA = TxtValueA.Text.Trim
                    'remove quote
                    sValueA = sValueA.Replace("'", "''")

                    If CboOperator.Text <> LM_OP_EQUAL Then
                        'Update 29/06/2010: CboOperator.Text = thì không Replace
                        'Tìm ký tự đặc biệt % thay bằng [%]
                        sValueA = sValueA.Replace("%", "[%]")
                        'Tìm ký tự đặc biệt * thay bằng [*]
                        sValueA = sValueA.Replace("*", "[*]")
                    End If

                    'Minh Hòa update 09/06/2009: Tìm kiếm theo tập hợp
                    Dim sChar As String = ";"
                    Dim sArr() As String = {""}
                    Dim sClauseT As String = ""
                    Dim sClauseTServer As String = ""
                    Dim bSplit As Boolean = False
                    If sValueA.Contains(sChar) Then
                        'Tìm theo tập hợp
                        sArr = Microsoft.VisualBasic.Split(sValueA, sChar)
                        bSplit = True
                    End If

                    Select Case CboOperator.Text
                        'Mặc định là Có chứa
                        Case LM_OP_CONTAIN
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        'sArr(i) = sArr(i).Replace("*", "[*]")
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Like '%" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like '%" & sArr(i) & "%'"
                                        End If

                                        If sClauseTServer = "" Then
                                            sClauseTServer &= "[" & cboField.Columns(1).Text & "]" & " Like N'%" & sArr(i) & "%'"
                                        Else
                                            sClauseTServer &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like N'%" & sArr(i) & "%'"
                                        End If

                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                                sClauseServer = "(" & sClauseTServer & ")"
                            Else
                                sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " Like '%" & sValueA & "%')"
                                sClauseServer = "( " & "[" & cboField.Columns(1).Text & "]" & " Like N'%" & sValueA & "%')"
                            End If

                            'Update 12/07/2010: incident 33798
                        Case LM_OP_NOTCONTAIN
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        'sArr(i) = sArr(i).Replace("*", "[*]")
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Not Like '%" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " And " & "[" & cboField.Columns(1).Text & "]" & " Not Like '%" & sArr(i) & "%'"
                                        End If

                                        If sClauseTServer = "" Then
                                            sClauseTServer &= "[" & cboField.Columns(1).Text & "]" & " Not Like N'%" & sArr(i) & "%'"
                                        Else
                                            sClauseTServer &= " And " & "[" & cboField.Columns(1).Text & "]" & " Not Like N'%" & sArr(i) & "%'"
                                        End If

                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                                sClauseServer = "(" & sClauseTServer & ")"
                            Else
                                sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " Not Like '%" & sValueA & "%')"
                                sClauseServer = "( " & "[" & cboField.Columns(1).Text & "]" & " Not Like N'%" & sValueA & "%')"
                            End If

                        Case LM_OP_STARTWITH
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Like '" & sArr(i) & "%'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like '" & sArr(i) & "%'"
                                        End If
                                        If sClauseTServer = "" Then
                                            sClauseTServer &= "[" & cboField.Columns(1).Text & "]" & " Like N'" & sArr(i) & "%'"
                                        Else
                                            sClauseTServer &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like N'" & sArr(i) & "%'"
                                        End If

                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                                sClauseServer = "(" & sClauseTServer & ")"
                            Else
                                sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " Like '" & sValueA & "%')"
                                sClauseServer = "( " & "[" & cboField.Columns(1).Text & "]" & " Like N'" & sValueA & "%')"
                            End If

                        Case LM_OP_ENDWITH
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Like '%" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like '%" & sArr(i) & "'"
                                        End If
                                        If sClauseTServer = "" Then
                                            sClauseTServer &= "[" & cboField.Columns(1).Text & "]" & " Like N'%" & sArr(i) & "'"
                                        Else
                                            sClauseTServer &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Like N'%" & sArr(i) & "'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                                sClauseServer = "(" & sClauseTServer & ")"
                            Else
                                sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " Like '%" & sValueA & "')"
                                sClauseServer = "(" & "[" & cboField.Columns(1).Text & "]" & " Like N'%" & sValueA & "')"
                            End If

                        Case LM_OP_EQUAL
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " = '" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " = '" & sArr(i) & "'"
                                        End If
                                        If sClauseTServer = "" Then
                                            sClauseTServer &= "[" & cboField.Columns(1).Text & "]" & " = N'" & sArr(i) & "'"
                                        Else
                                            sClauseTServer &= " Or " & "[" & cboField.Columns(1).Text & "]" & " = N'" & sArr(i) & "'"
                                        End If

                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                                sClauseServer = "(" & sClauseTServer & ")"
                            Else
                                sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " = '" & sValueA & "')"
                                sClauseServer = "(" & "[" & cboField.Columns(1).Text & "]" & " = N'" & sValueA & "')"
                                If sValueA = "" Then sClause = "(" & sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & " is NULL)"
                            End If

                        Case LM_OP_NOTEQUAL
                            If bSplit Then 'Tìm theo tập hợp
                                For i As Integer = 0 To sArr.GetUpperBound(0)
                                    If sArr(i) <> "" Then
                                        If sClauseT = "" Then
                                            sClauseT &= "[" & cboField.Columns(1).Text & "]" & " Not Like '" & sArr(i) & "'"
                                        Else
                                            sClauseT &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Not Like '" & sArr(i) & "'"
                                        End If
                                        If sClauseTServer = "" Then
                                            sClauseTServer &= "[" & cboField.Columns(1).Text & "]" & " Not Like N'" & sArr(i) & "'"
                                        Else
                                            sClauseTServer &= " Or " & "[" & cboField.Columns(1).Text & "]" & " Not Like N'" & sArr(i) & "'"
                                        End If
                                    End If
                                Next
                                sClause = "(" & sClauseT & ")"
                                sClauseServer = "(" & sClauseTServer & ")"
                            Else
                                sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " Not Like '" & sValueA & "')"
                                sClauseServer = "(" & "[" & cboField.Columns(1).Text & "]" & " Not Like N'" & sValueA & "')"

                                If sValueA = "" Then sClause = "(" & sClause & " And " & "[" & cboField.Columns(1).Text & "]" & " is not NULL)"
                            End If

                    End Select
                End If

            Case FinderTypeEnum.lmFinderDate
                ' Date: 05/12/2007 Thay thế các hàm SQLDateSave  thành DateSave
                'Mặc định là Trong khoảng
                Select Case CboOperator.Text
                    Case LM_OP_BETWEEN
                        c1dateB.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >= #" & DateSave(c1dateA.Value) & "#"
                        sClause = sClause & " And " & "[" & cboField.Columns(1).Text & "]" & " < #" & DateSave(CDate(c1dateB.Value).AddDays(1)) & "#)"

                        sClauseServer = "(" & cboField.Columns(1).Text & " >= " & SQLDateSave(c1dateA.Value)
                        sClauseServer = sClauseServer & " And " & cboField.Columns(1).Text & " - 1 < " & SQLDateSave(c1dateB.Value) & ")"

                    Case LM_OP_EQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >= #" & DateSave(c1dateA.Value) & "#"
                        sClause = sClause & " And " & "[" & cboField.Columns(1).Text & "]" & "  < #" & DateSave(CDate(c1dateA.Value).AddDays(1)) & "# )"

                        sClauseServer = "(" & cboField.Columns(1).Text & " >= " & SQLDateSave(c1dateA.Value)
                        sClauseServer = sClauseServer & " And " & cboField.Columns(1).Text & " - 1 < " & SQLDateSave(c1dateA.Value) & ")"

                    Case LM_OP_GREATEREQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >= #" & DateSave(c1dateA.Value) & "#)"
                        sClauseServer = "(" & cboField.Columns(1).Text & " >= " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_GREATER
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >= " & " #" & DateSave(CDate(c1dateA.Value).AddDays(1)) & "#)"
                        sClauseServer = "(" & cboField.Columns(1).Text & " - 1 >= " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_LESSEQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " <= #" & DateSave(c1dateA.Value) & "#)"
                        sClauseServer = "(" & cboField.Columns(1).Text & " <= " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_LESS
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " < #" & DateSave(c1dateA.Value) & "#)"
                        sClauseServer = "(" & cboField.Columns(1).Text & " < " & SQLDateSave(c1dateA.Value) & ")"
                    Case LM_OP_NOTBETWEEN
                        c1dateB.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " < #" & DateSave(c1dateA.Value) & "#"
                        sClause = sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & "  >= #" & DateSave(CDate(c1dateB.Value).AddDays(1)) & "#)"

                        sClauseServer = "(" & cboField.Columns(1).Text & " < " & SQLDateSave(c1dateA.Value)
                        sClauseServer = sClauseServer & " Or " & cboField.Columns(1).Text & " - 1 >= " & SQLDateSave(c1dateB.Value) & ")"

                    Case LM_OP_NOTEQUAL
                        c1dateA.UpdateValueWithCurrentText()
                        sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " < #" & DateSave(c1dateA.Value) & "#"
                        sClause = sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & "  >= #" & DateSave(CDate(c1dateA.Value).AddDays(1)) & "# )"

                        sClauseServer = "(" & cboField.Columns(1).Text & " < " & SQLDateSave(c1dateA.Value)
                        sClauseServer = sClauseServer & " Or " & cboField.Columns(1).Text & " - 1 >= " & SQLDateSave(c1dateA.Value) & ")"

                End Select

            Case FinderTypeEnum.lmFinderTinyInt 'Update 02/07/2010: Checkbox
                sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " = " & SQLNumber(chk.Checked) & ")"
                sClauseServer = sClause
            Case FinderTypeEnum.lmFinderInt, FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                Select Case CboOperator.Text
                    'Mặc định là bằng
                    Case LM_OP_EQUAL
                        If IsNumeric(TxtValueA.Text) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " = " & CDbl(TxtValueA.Text) & ")"
                            sClauseServer = "( " & cboField.Columns(1).Text & " = " & TxtValueA.Text & ")"
                        End If
                        If Number(TxtValueA.Text) = 0 And TxtValueA.Text <> "" Then sClause = "(" & sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & " is NULL )"
                    Case LM_OP_GREATEREQUAL
                        If IsNumeric(TxtValueA.Text) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " >= " & CDbl(TxtValueA.Text) & ")"
                            sClauseServer = "( " & cboField.Columns(1).Text & " >= " & TxtValueA.Text & ")"
                        End If
                    Case LM_OP_GREATER
                        If IsNumeric(TxtValueA.Text) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " > " & CDbl(TxtValueA.Text) & ")"
                            sClauseServer = "( " & cboField.Columns(1).Text & " > " & TxtValueA.Text & ")"
                        End If
                    Case LM_OP_LESSEQUAL
                        If IsNumeric(TxtValueA.Text) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " <= " & CDbl(TxtValueA.Text) & ")"
                            sClauseServer = "( " & cboField.Columns(1).Text & " <= " & TxtValueA.Text & ")"
                        End If
                    Case LM_OP_LESS
                        If IsNumeric(TxtValueA.Text) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " < " & CDbl(TxtValueA.Text) & ")"
                            sClauseServer = "( " & cboField.Columns(1).Text & " < " & TxtValueA.Text & ")"
                        End If
                    Case LM_OP_BETWEEN
                        If IsNumeric(TxtValueA.Text) And IsNumeric(TxtValueB.Text) Then
                            If TxtValueB.Text = "" Then TxtValueB.Text = TxtValueA.Text
                            sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " >=" & CDbl(TxtValueA.Text)
                            sClause = sClause & " And " & "[" & cboField.Columns(1).Text & "]" & " <= " & CDbl(TxtValueB.Text) & ")"

                            sClauseServer = "( " & cboField.Columns(1).Text & " between " & TxtValueA.Text
                            sClauseServer = sClauseServer & " and " & TxtValueB.Text & ")"

                        End If
                    Case LM_OP_NOTBETWEEN
                        If IsNumeric(TxtValueA.Text) And IsNumeric(TxtValueB.Text) Then
                            sClause = "(" & "[" & cboField.Columns(1).Text & "]" & " <" & CDbl(TxtValueA.Text)
                            sClause = sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & " > " & CDbl(TxtValueB.Text) & ")"

                            sClauseServer = "( " & cboField.Columns(1).Text & " not between " & TxtValueA.Text
                            sClauseServer = sClauseServer & " and " & TxtValueB.Text & ")"

                        End If
                    Case LM_OP_NOTEQUAL
                        If IsNumeric(TxtValueA.Text) Then
                            sClause = "( " & "[" & cboField.Columns(1).Text & "]" & " <> " & CDbl(TxtValueA.Text) & ")"
                            sClauseServer = "( " & cboField.Columns(1).Text & " <> " & TxtValueA.Text & ")"
                        End If
                        If Number(TxtValueA.Text) = 0 And TxtValueA.Text <> "" Then sClause = "(" & sClause & " Or " & "[" & cboField.Columns(1).Text & "]" & " is not NULL )"
                End Select
        End Select

        sClause = sClause.Trim
        sClauseServer = sClauseServer.Trim

        If sClause <> "" Then
            If strSQLAdvanced <> "" Then
                If Label.Text = LM_AND Then
                    strSQLAdvanced = strSQLAdvanced & " AND " & sClause
                    strSQLAdvancedForReport = strSQLAdvancedForReport & " AND " & sClauseServer
                Else
                    strSQLAdvanced = strSQLAdvanced & " OR " & sClause
                    strSQLAdvancedForReport = strSQLAdvancedForReport & " OR " & sClauseServer
                End If
            Else
                strSQLAdvanced = sClause
                strSQLAdvancedForReport = sClauseServer
            End If
        End If

    End Sub

    Private Sub ButFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButFind.Click
        Try
            Dim i As Integer

            strSQLAdvanced = ""
            'Lưu lại giá trị tìm kiếm vào mảng
            CheckDisplayValueAddList()

            '*********************
            'Update 10/03/2009
            'Tăng TemplateID để lưu dữ liệu tìm kiếm 
            Dim sSQLInsert As New System.Text.StringBuilder
            If iMaxTemplateID < 10 Then
                iMaxTemplateID += 1
                iLastTemplateID = iMaxTemplateID
            Else
                iLastTemplateID += 1
            End If
            If iLastTemplateID > 10 Then
                iLastTemplateID = 1
            End If
            iTemplateID = iLastTemplateID

            sSQLInsert.AppendLine(SQLDeleteD91T1100())

            Dim sSQL As String = "Insert Into D91T1100(FormID,Mode,SubMode,UserID,TemplateID,FieldOrderNo,FieldName,FieldNameU, " _
                    & "StrCompare, StrValue, StrValueU, DateFrom, DateTo, NumFrom, NumTo, CreateDate, Operator) " _
                    & "Values(" & SQLString(_formID) & "," & SQLString(_mode) & ",''," & SQLString(gsUserID) & "," & iTemplateID.ToString & ",{FieldOrderNo}, {FieldName}, {FieldNameU},{StrCompare}, " _
                    & "{StrValue}, {StrValueU}, {DateFrom},{DateTo},{NumFrom},{NumTo},getDate(),{Operator})"

            Dim sSQLTemp As String = ""
            If LM_FindClient = 2 Then ' Update 07/10/2010: Tìm kiếm Client nhưng trả về giá trị dạng server
                For i = 1 To idx
                    'sSQLTemp = sSQL
                    Select Case i
                        Case 1
                            ListFindClientServer(CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, Label1, chk01)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, chk01)
                        Case 2
                            ListFindClientServer(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, Label2, chk02)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, chk02, Label2.Text)
                        Case 3
                            ListFindClientServer(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, Label3, chk03)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, chk03, Label3.Text)
                        Case 4
                            ListFindClientServer(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, Label4, chk04)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, chk04, Label4.Text)
                        Case 5
                            ListFindClientServer(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, Label5, chk05)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, chk05, Label5.Text)
                        Case 6
                            ListFindClientServer(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, Label6, chk06)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, chk06, Label6.Text)
                        Case 7
                            ListFindClientServer(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, Label7, chk07)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, chk07, Label7.Text)
                        Case 8
                            ListFindClientServer(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, Label8, chk08)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, chk08, Label8.Text)
                        Case 9
                            ListFindClientServer(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, Label9, chk09)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, chk09, Label9.Text)
                        Case 10
                            ListFindClientServer(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, Label10, chk10)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, chk10, Label10.Text)
                    End Select
                    sSQLInsert.AppendLine(sSQLTemp)
                    sSQLTemp = ""
                Next i

            ElseIf LM_FindClient = 1 Then 'Tìm kiếm Client
                For i = 1 To idx
                    'sSQLTemp = sSQL
                    Select Case i
                        Case 1
                            ListFindClient(CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, Label1, chk01)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, chk01)
                        Case 2
                            ListFindClient(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, Label2, chk02)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, chk02, Label2.Text)
                        Case 3
                            ListFindClient(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, Label3, chk03)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, chk03, Label3.Text)
                        Case 4
                            ListFindClient(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, Label4, chk04)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, chk04, Label4.Text)
                        Case 5
                            ListFindClient(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, Label5, chk05)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, chk05, Label5.Text)
                        Case 6
                            ListFindClient(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, Label6, chk06)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, chk06, Label6.Text)
                        Case 7
                            ListFindClient(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, Label7, chk07)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, chk07, Label7.Text)
                        Case 8
                            ListFindClient(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, Label8, chk08)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, chk08, Label8.Text)
                        Case 9
                            ListFindClient(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, Label9, chk09)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, chk09, Label9.Text)
                        Case 10
                            ListFindClient(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, Label10, chk10)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, chk10, Label10.Text)
                    End Select
                    sSQLInsert.AppendLine(sSQLTemp)
                    sSQLTemp = ""
                Next i
            Else ' Tìm kiếm trên Server
                For i = 1 To idx
                    'sSQLTemp = sSQL
                    Select Case i
                        Case 1
                            ListFind(CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, Label1, chk01)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, chk01)
                        Case 2
                            ListFind(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, Label2, chk02)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, chk02, Label2.Text)
                        Case 3
                            ListFind(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, Label3, chk03)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, chk03, Label3.Text)
                        Case 4
                            ListFind(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, Label4, chk04)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, chk04, Label4.Text)
                        Case 5
                            ListFind(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, Label5, chk05)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, chk05, Label5.Text)
                        Case 6
                            ListFind(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, Label6, chk06)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, chk06, Label6.Text)
                        Case 7
                            ListFind(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, Label7, chk07)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, chk07, Label7.Text)
                        Case 8
                            ListFind(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, Label8, chk08)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, chk08, Label8.Text)
                        Case 9
                            ListFind(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, Label9, chk09)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, chk09, Label9.Text)
                        Case 10
                            ListFind(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, Label10, chk10)
                            sSQLTemp = ReplaceSQLInsert(sSQL, i.ToString, CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, chk10, Label10.Text)
                    End Select
                    sSQLInsert.AppendLine(sSQLTemp)
                    sSQLTemp = ""
                Next i
            End If

            'Check order by
            Dim strOrder As String
            Dim strComma As String
            strOrder = ""
            strComma = ""

            If blnOrder Then
                If Not xSortOrder Is Nothing Then
                    For i = 0 To UBound(xSortOrder)
                        strComma = IIf(Trim(strOrder) = "", "", " , ").ToString
                        strOrder = strOrder & strComma & Trim(xSortOrder(i, 1)) & " " & IIf(xSortOrder(i, 2) = LM_ASC, "", " DESC ").ToString
                    Next i
                End If

                If Trim(strSQLAdvanced) <> "" And Trim(strOrder) <> "" Then
                    strSQLAdvanced = strSQLAdvanced & " ORDER BY " & strOrder
                End If
            End If


            'Thực thi câu SQL luu du lieu xuong bang D91T1100
            ExecuteSQL(sSQLInsert.ToString)

            Me.Close()

        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try

    End Sub

    Private Sub CheckDisplayValueAddList()
        If TxtValueA01.Visible Then
            AddValueInList(TxtValueA01)
            If TxtValueB01.Visible Then
                AddValueInList(TxtValueB01)
            End If
        End If

        If TxtValueA02.Visible Then
            AddValueInList(TxtValueA02)
            If TxtValueB02.Visible Then
                AddValueInList(TxtValueB02)
            End If
        End If

        If TxtValueA03.Visible Then
            AddValueInList(TxtValueA03)
            If TxtValueB03.Visible Then
                AddValueInList(TxtValueB03)
            End If
        End If

        If TxtValueA04.Visible Then
            AddValueInList(TxtValueA04)
            If TxtValueB04.Visible Then
                AddValueInList(TxtValueB04)
            End If
        End If

        If TxtValueA05.Visible Then
            AddValueInList(TxtValueA05)
            If TxtValueB05.Visible Then
                AddValueInList(TxtValueB05)
            End If
        End If

        If TxtValueA06.Visible Then
            AddValueInList(TxtValueA06)
            If TxtValueB06.Visible Then
                AddValueInList(TxtValueB06)
            End If
        End If

        If TxtValueA07.Visible Then
            AddValueInList(TxtValueA07)
            If TxtValueB07.Visible Then
                AddValueInList(TxtValueB07)
            End If
        End If

        If TxtValueA08.Visible Then
            AddValueInList(TxtValueA08)
            If TxtValueB08.Visible Then
                AddValueInList(TxtValueB08)
            End If
        End If

        If TxtValueA09.Visible Then
            AddValueInList(TxtValueA09)
            If TxtValueB09.Visible Then
                AddValueInList(TxtValueB09)
            End If
        End If

        If TxtValueA10.Visible Then
            AddValueInList(TxtValueA10)
            If TxtValueB10.Visible Then
                AddValueInList(TxtValueB10)
            End If
        End If
    End Sub

    Private Sub AddValueInList(ByVal txtValue As System.Windows.Forms.TextBox)
        If lsValueListFind Is Nothing Then
            lsValueListFind = New List(Of String)
            lsValueListFind.Add("")
            lsValueListFind.Add(txtValue.Text)
        Else
            Select Case txtValue.Name
                Case TxtValueA01.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA01) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If
                Case TxtValueA02.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA02) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If
                Case TxtValueA03.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA03) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If

                Case TxtValueA04.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA04) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If

                Case TxtValueA05.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA05) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If

                Case TxtValueA06.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA06) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If

                Case TxtValueA07.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA07) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If

                Case TxtValueA08.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA08) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If

                Case TxtValueA09.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA09) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If

                Case TxtValueA10.Name
                    If Not lsValueListFind.Exists(AddressOf EqualsValueA10) Then
                        lsValueListFind.Add(txtValue.Text)
                    End If
            End Select


        End If

    End Sub

    Private Function EqualsValueA02(ByVal s As String) _
        As Boolean
        If s = TxtValueA02.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA01(ByVal s As String) _
        As Boolean
        If s = TxtValueA01.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA03(ByVal s As String) _
        As Boolean
        If s = TxtValueA03.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA04(ByVal s As String) _
        As Boolean
        If s = TxtValueA04.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA05(ByVal s As String) _
        As Boolean
        If s = TxtValueA05.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA06(ByVal s As String) _
        As Boolean
        If s = TxtValueA06.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA07(ByVal s As String) _
        As Boolean
        If s = TxtValueA07.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA08(ByVal s As String) _
        As Boolean
        If s = TxtValueA08.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA09(ByVal s As String) _
        As Boolean
        If s = TxtValueA09.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function EqualsValueA10(ByVal s As String) _
        As Boolean
        If s = TxtValueA10.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub ButOr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButOr.Click

        AddRow(LM_OR)
        EnableButton(False)
        Center(Me)
    End Sub

    Private Sub GetTemplate()
        'Gan Template tim kiem
        'TxtValueA01.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        'TxtValueA01.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        'TxtValueA01.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper

        TxtValueA01.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA01.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA02.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA02.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA03.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA03.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA04.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA04.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA05.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA05.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA06.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA06.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA07.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA07.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA08.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA08.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA09.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA09.AutoCompleteSource = AutoCompleteSource.CustomSource
        TxtValueA10.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TxtValueA10.AutoCompleteSource = AutoCompleteSource.CustomSource

        If lsValueListFind IsNot Nothing Then
            TxtValueA01.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA02.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA03.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA04.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA05.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA06.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA07.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA08.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA09.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
            TxtValueA10.AutoCompleteCustomSource.AddRange(lsValueListFind.ToArray)
        End If
    End Sub

    Private Sub D99F0006_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LoadLanguage()

          

            idx = 1
            Label1.Top = 9
            Label1.Left = CboField01.Left + LM_SPACE
            CboOperator01.Enabled = False

            TxtValueA01.Text = ""
            TxtValueA01.Left = CboOperator01.Left + CboOperator01.Width + LM_SPACE
            TxtValueA01.Width = LM_DOUBLE_WIDTH + LM_SPACE
            TxtValueA01.Visible = True
            TxtValueA01.Enabled = False

            '***********************************************************
            'Update 19/10/2010: Tạm thời bỏ hàm GetTemplate do bị lỗi Font theo Incident 34234
            'Gan Template tim kiem
            'GetTemplate()

            '***********************************************************

            TxtValueB01.Text = ""
            TxtValueB01.Left = TxtValueA01.Left + TxtValueA01.Width + LM_SPACE
            TxtValueB01.Visible = False

            c1dateA01.Value = Now
            c1dateA01.Left = CboOperator01.Left + CboOperator01.Width + LM_SPACE
            c1dateA01.Visible = False

            c1dateB01.Value = Now
            c1dateB01.Left = c1dateA01.Left + c1dateA01.Width + LM_SPACE
            c1dateB01.Visible = False

            chk01.Left = CboOperator01.Left + CboOperator01.Width + LM_SPACE
            chk01.Visible = False

            '***********************************************************
            ButBack.Enabled = False
            btnBackFirst.Enabled = False
            StuffFieldNameCombo()

            VisibleControl()

            'Set  TabIndex for Controls
            SetTabIndexControl()

            EnableButton(False)
            'Minh Hoa update 25/08/2008: lấy giá trị mặc định của dòng đầu tiên
            CboField01.Text = CboField01.Columns(0).Text
            CboField01_Close(Nothing, Nothing)

            '*********************
            'Update 10/03/2009
            LoadTemplateID()
            '*********************

            'Kiểm tra có sử dụng Unicode ko 
            InputbyUnicode(Me, _useUnicode)
            'Set độ phân giải theo Window
            SetResolutionForm(Me)

        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try

    End Sub

    Private Sub SetTabIndexControl()
        GroupBox1.TabIndex = 0

        CboField01.TabIndex = 1
        CboOperator01.TabIndex = CboField01.TabIndex + 1
        TxtValueA01.TabIndex = CboOperator01.TabIndex + 1
        TxtValueB01.TabIndex = TxtValueA01.TabIndex + 1
        c1dateA01.TabIndex = TxtValueB01.TabIndex + 1
        c1dateB01.TabIndex = c1dateA01.TabIndex + 1
        chk01.TabIndex = c1dateB01.TabIndex + 1

        CboField02.TabIndex = chk01.TabIndex + 1
        CboOperator02.TabIndex = CboField02.TabIndex + 1
        TxtValueA02.TabIndex = CboOperator02.TabIndex + 1
        TxtValueB02.TabIndex = TxtValueA02.TabIndex + 1
        c1dateA02.TabIndex = TxtValueB02.TabIndex + 1
        c1dateB02.TabIndex = c1dateA02.TabIndex + 1
        chk02.TabIndex = c1dateB02.TabIndex + 1

        CboField03.TabIndex = chk02.TabIndex + 1
        CboOperator03.TabIndex = CboField03.TabIndex + 1
        TxtValueA03.TabIndex = CboOperator03.TabIndex + 1
        TxtValueB03.TabIndex = TxtValueA03.TabIndex + 1
        c1dateA03.TabIndex = TxtValueB03.TabIndex + 1
        c1dateB03.TabIndex = c1dateA03.TabIndex + 1
        chk03.TabIndex = c1dateB03.TabIndex + 1

        CboField04.TabIndex = chk03.TabIndex + 1
        CboOperator04.TabIndex = CboField04.TabIndex + 1
        TxtValueA04.TabIndex = CboOperator04.TabIndex + 1
        TxtValueB04.TabIndex = TxtValueA04.TabIndex + 1
        c1dateA04.TabIndex = TxtValueB04.TabIndex + 1
        c1dateB04.TabIndex = c1dateA04.TabIndex + 1
        chk04.TabIndex = c1dateB04.TabIndex + 1

        CboField05.TabIndex = chk04.TabIndex + 1
        CboOperator05.TabIndex = CboField05.TabIndex + 1
        TxtValueA05.TabIndex = CboOperator05.TabIndex + 1
        TxtValueB05.TabIndex = TxtValueA05.TabIndex + 1
        c1dateA05.TabIndex = TxtValueB05.TabIndex + 1
        c1dateB05.TabIndex = c1dateA05.TabIndex + 1
        chk05.TabIndex = c1dateB05.TabIndex + 1

        CboField06.TabIndex = chk05.TabIndex + 1
        CboOperator06.TabIndex = CboField06.TabIndex + 1
        TxtValueA06.TabIndex = CboOperator06.TabIndex + 1
        TxtValueB06.TabIndex = TxtValueA06.TabIndex + 1
        c1dateA06.TabIndex = TxtValueB06.TabIndex + 1
        c1dateB06.TabIndex = c1dateA06.TabIndex + 1
        chk06.TabIndex = c1dateB06.TabIndex + 1

        CboField07.TabIndex = chk06.TabIndex + 1
        CboOperator07.TabIndex = CboField07.TabIndex + 1
        TxtValueA07.TabIndex = CboOperator07.TabIndex + 1
        TxtValueB07.TabIndex = TxtValueA07.TabIndex + 1
        c1dateA07.TabIndex = TxtValueB07.TabIndex + 1
        c1dateB07.TabIndex = c1dateA07.TabIndex + 1
        chk07.TabIndex = c1dateB07.TabIndex + 1

        CboField08.TabIndex = chk07.TabIndex + 1
        CboOperator08.TabIndex = CboField08.TabIndex + 1
        TxtValueA08.TabIndex = CboOperator08.TabIndex + 1
        TxtValueB08.TabIndex = TxtValueA08.TabIndex + 1
        c1dateA08.TabIndex = TxtValueB08.TabIndex + 1
        c1dateB08.TabIndex = c1dateA08.TabIndex + 1
        chk08.TabIndex = c1dateB08.TabIndex + 1

        CboField09.TabIndex = chk08.TabIndex + 1
        CboOperator09.TabIndex = CboField09.TabIndex + 1
        TxtValueA09.TabIndex = CboOperator09.TabIndex + 1
        TxtValueB09.TabIndex = TxtValueA09.TabIndex + 1
        c1dateA09.TabIndex = TxtValueB09.TabIndex + 1
        c1dateB09.TabIndex = c1dateA09.TabIndex + 1
        chk09.TabIndex = c1dateB09.TabIndex + 1

        CboField10.TabIndex = chk09.TabIndex + 1
        CboOperator10.TabIndex = CboField10.TabIndex + 1
        TxtValueA10.TabIndex = CboOperator10.TabIndex + 1
        TxtValueB10.TabIndex = TxtValueA10.TabIndex + 1
        c1dateA10.TabIndex = TxtValueB10.TabIndex + 1
        c1dateB10.TabIndex = c1dateA10.TabIndex + 1
        chk10.TabIndex = c1dateB10.TabIndex + 1

        ButFind.TabIndex = chk10.TabIndex + 1 '61
        ButClose.TabIndex = ButFind.TabIndex + 1
        ButAnd.TabIndex = ButClose.TabIndex + 1
        ButOr.TabIndex = ButAnd.TabIndex + 1
        ButBack.TabIndex = ButOr.TabIndex + 1
        'ButHelp.TabIndex = 66
        'ButClose.TabIndex = 67

    End Sub
    Private Sub VisibleControl()
        CboField02.Visible = False
        CboField03.Visible = False
        CboField04.Visible = False
        CboField05.Visible = False
        CboField06.Visible = False
        CboField07.Visible = False
        CboField08.Visible = False
        CboField09.Visible = False
        CboField10.Visible = False

        CboOperator02.Visible = False
        CboOperator03.Visible = False
        CboOperator04.Visible = False
        CboOperator05.Visible = False
        CboOperator06.Visible = False
        CboOperator07.Visible = False
        CboOperator08.Visible = False
        CboOperator09.Visible = False
        CboOperator10.Visible = False

        TxtValueA02.Visible = False
        TxtValueA03.Visible = False
        TxtValueA04.Visible = False
        TxtValueA05.Visible = False
        TxtValueA06.Visible = False
        TxtValueA07.Visible = False
        TxtValueA08.Visible = False
        TxtValueA09.Visible = False
        TxtValueA10.Visible = False

        TxtValueB02.Visible = False
        TxtValueB03.Visible = False
        TxtValueB04.Visible = False
        TxtValueB05.Visible = False
        TxtValueB06.Visible = False
        TxtValueB07.Visible = False
        TxtValueB08.Visible = False
        TxtValueB09.Visible = False
        TxtValueB10.Visible = False

        c1dateA02.Visible = False
        c1dateA03.Visible = False
        c1dateA04.Visible = False
        c1dateA05.Visible = False
        c1dateA06.Visible = False
        c1dateA07.Visible = False
        c1dateA08.Visible = False
        c1dateA09.Visible = False
        c1dateA10.Visible = False

        c1dateB02.Visible = False
        c1dateB03.Visible = False
        c1dateB04.Visible = False
        c1dateB05.Visible = False
        c1dateB06.Visible = False
        c1dateB07.Visible = False
        c1dateB08.Visible = False
        c1dateB09.Visible = False
        c1dateB10.Visible = False


        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Label10.Visible = False

        chk02.Visible = False
        chk03.Visible = False
        chk04.Visible = False
        chk05.Visible = False
        chk06.Visible = False
        chk07.Visible = False
        chk08.Visible = False
        chk09.Visible = False
        chk10.Visible = False

    End Sub

    Private Sub LoadLanguage()
        btnPrevious.Text = r("_Tro_lai") & " (F8)"
        If geLanguage = EnumLanguage.Vietnamese Then
            sMsgNhapSoQuaLon = "Nhập số quá lớn"
            sMsgNhapSoChuaDung = "Nhập số chưa đúng"
        Else
            sMsgNhapSoQuaLon = "Too large number"
            sMsgNhapSoChuaDung = "Invalid number"
        End If
    End Sub

    Private Sub AddRow(ByVal sOperator As String)
        Try
            idx = idx + 1

            ButBack.Enabled = True
            btnBackFirst.Enabled = True
            Me.Height = Me.Height + LM_H
            GroupBox1.Height = GroupBox1.Height + LM_H
            ButAnd.Top = ButAnd.Top + LM_H
            ButOr.Top = ButOr.Top + LM_H
            ButBack.Top = ButBack.Top + LM_H
            ButFind.Top = ButFind.Top + LM_H
            'ButHelp.Top = ButHelp.Top + LM_H
            ButClose.Top = ButClose.Top + LM_H
            btnBackFirst.Top = btnBackFirst.Top + LM_H
            btnPrevious.Top = btnPrevious.Top + LM_H
            Select Case idx
                Case 2

                    'Display Label2
                    Label2.Visible = True
                    Label2.Top = Label1.Top + LM_H
                    Label2.Left = Label1.Left
                    Label2.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label2.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label2.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label2.Visible = True

                    'Display cboField2
                    CboField02.Visible = True
                    CboField02.Top = CboField01.Top + LM_H
                    CboField02.Left = CboField01.Left
                    CboField02.Visible = True

                    'Display cboOperator2
                    CboOperator02.Visible = True
                    CboOperator02.Top = CboOperator01.Top + LM_H
                    CboOperator02.Left = CboOperator01.Left
                    CboOperator02.Enabled = False

                    'Display txtValueA2
                    TxtValueA02.Visible = True
                    TxtValueA02.Enabled = False
                    TxtValueA02.Top = TxtValueA01.Top + LM_H
                    TxtValueA02.Left = TxtValueA01.Left
                    TxtValueA02.Text = ""

                    'Display txtValueB2
                    'TxtValueB02.Visible = True
                    TxtValueB02.Top = TxtValueB01.Top + LM_H
                    TxtValueB02.Left = TxtValueA01.Left
                    TxtValueB02.Text = ""

                    'Display c1dateA02
                    'c1dateA02.Visible = True
                    c1dateA02.Top = c1dateA01.Top + LM_H
                    c1dateA02.Left = c1dateA01.Left
                    c1dateA02.Value = Now

                    'Display c1dateB01
                    'c1dateB02.Visible = True
                    c1dateB02.Top = c1dateB01.Top + LM_H
                    c1dateB02.Left = c1dateB01.Left
                    c1dateB02.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk02.Top = chk01.Top + LM_H
                    chk02.Left = chk01.Left
                    chk02.Checked = False

                    CboField02.Focus()
                Case 3

                    'Display Label3
                    Label3.Visible = True
                    Label3.Top = Label2.Top + LM_H
                    Label3.Left = Label2.Left
                    Label3.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label3.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label3.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label3.Visible = True

                    'Display cboField3
                    CboField03.Visible = True
                    CboField03.Top = CboField02.Top + LM_H
                    CboField03.Left = CboField02.Left
                    CboField03.Visible = True

                    'Display cboOperator3
                    CboOperator03.Visible = True
                    CboOperator03.Top = CboOperator02.Top + LM_H
                    CboOperator03.Left = CboOperator02.Left
                    CboOperator03.Enabled = False

                    'Display txtValueA3
                    TxtValueA03.Visible = True
                    TxtValueA03.Enabled = False
                    TxtValueA03.Top = TxtValueA02.Top + LM_H
                    TxtValueA03.Left = TxtValueA02.Left
                    TxtValueA03.Text = ""

                    'Display txtValueB3
                    'TxtValueB03.Visible = True
                    TxtValueB03.Top = TxtValueB02.Top + LM_H
                    TxtValueB03.Left = TxtValueA01.Left
                    TxtValueB03.Text = ""

                    'Display c1dateA03
                    'c1dateA03.Visible = True
                    c1dateA03.Top = c1dateA02.Top + LM_H
                    c1dateA03.Left = c1dateA02.Left
                    c1dateA03.Value = Now

                    'Display c1dateB03
                    'c1dateB03.Visible = True
                    c1dateB03.Top = c1dateB02.Top + LM_H
                    c1dateB03.Left = c1dateB02.Left
                    c1dateB03.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk03.Top = chk02.Top + LM_H
                    chk03.Left = chk01.Left
                    chk03.Checked = False


                    CboField03.Focus()
                Case 4

                    'Display Label4
                    Label4.Visible = True
                    Label4.Top = Label3.Top + LM_H
                    Label4.Left = Label3.Left
                    Label4.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label4.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label4.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label4.Visible = True

                    'Display cboField4
                    CboField04.Visible = True
                    CboField04.Top = CboField03.Top + LM_H
                    CboField04.Left = CboField03.Left
                    CboField04.Visible = True

                    'Display cboOperator4
                    CboOperator04.Visible = True
                    CboOperator04.Top = CboOperator03.Top + LM_H
                    CboOperator04.Left = CboOperator03.Left
                    CboOperator04.Enabled = False

                    'Display txtValueA4
                    TxtValueA04.Visible = True
                    TxtValueA04.Enabled = False
                    TxtValueA04.Top = TxtValueA03.Top + LM_H
                    TxtValueA04.Left = TxtValueA03.Left
                    TxtValueA04.Text = ""

                    'Display txtValueB4
                    'TxtValueB04.Visible = True
                    TxtValueB04.Top = TxtValueB03.Top + LM_H
                    TxtValueB04.Left = TxtValueA01.Left
                    TxtValueB04.Text = ""

                    'Display c1dateA04
                    'c1dateA04.Visible = True
                    c1dateA04.Top = c1dateA03.Top + LM_H
                    c1dateA04.Left = c1dateA03.Left
                    c1dateA04.Value = Now

                    'Display c1dateB04
                    'c1dateB04.Visible = True
                    c1dateB04.Top = c1dateB03.Top + LM_H
                    c1dateB04.Left = c1dateB03.Left
                    c1dateB04.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk04.Top = chk03.Top + LM_H
                    chk04.Left = chk01.Left
                    chk04.Checked = False


                    CboField04.Focus()
                Case 5

                    'Display Label5
                    Label5.Visible = True
                    Label5.Top = Label4.Top + LM_H
                    Label5.Left = Label4.Left
                    Label5.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label5.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label5.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label5.Visible = True

                    'Display cboField5
                    CboField05.Visible = True
                    CboField05.Top = CboField04.Top + LM_H
                    CboField05.Left = CboField04.Left
                    CboField05.Visible = True

                    'Display cboOperator5
                    CboOperator05.Visible = True
                    CboOperator05.Top = CboOperator04.Top + LM_H
                    CboOperator05.Left = CboOperator04.Left
                    CboOperator05.Enabled = False

                    'Display txtValueA5
                    TxtValueA05.Visible = True
                    TxtValueA05.Enabled = False
                    TxtValueA05.Top = TxtValueA04.Top + LM_H
                    TxtValueA05.Left = TxtValueA04.Left
                    TxtValueA05.Text = ""

                    'Display txtValueB5
                    'TxtValueB05.Visible = True
                    TxtValueB05.Top = TxtValueB04.Top + LM_H
                    TxtValueB05.Left = TxtValueA01.Left
                    TxtValueB05.Text = ""

                    'Display c1dateA05
                    'c1dateA05.Visible = True
                    c1dateA05.Top = c1dateA04.Top + LM_H
                    c1dateA05.Left = c1dateA04.Left
                    c1dateA05.Value = Now

                    'Display c1dateB05
                    'c1dateB05.Visible = True
                    c1dateB05.Top = c1dateB04.Top + LM_H
                    c1dateB05.Left = c1dateB04.Left
                    c1dateB05.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk05.Top = chk04.Top + LM_H
                    chk05.Left = chk01.Left
                    chk05.Checked = False

                    CboField05.Focus()
                Case 6

                    'Display Label6
                    Label6.Visible = True
                    Label6.Top = Label5.Top + LM_H
                    Label6.Left = Label5.Left
                    Label6.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label6.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label6.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label6.Visible = True

                    'Display cboField6
                    CboField06.Visible = True
                    CboField06.Top = CboField05.Top + LM_H
                    CboField06.Left = CboField05.Left
                    CboField06.Visible = True

                    'Display cboOperator6
                    CboOperator06.Visible = True
                    CboOperator06.Top = CboOperator05.Top + LM_H
                    CboOperator06.Left = CboOperator05.Left
                    CboOperator06.Enabled = False

                    'Display txtValueA6
                    TxtValueA06.Visible = True
                    TxtValueA06.Enabled = False
                    TxtValueA06.Top = TxtValueA05.Top + LM_H
                    TxtValueA06.Left = TxtValueA05.Left
                    TxtValueA06.Text = ""

                    'Display txtValueB6
                    'TxtValueB06.Visible = True
                    TxtValueB06.Top = TxtValueB05.Top + LM_H
                    TxtValueB06.Left = TxtValueA01.Left
                    TxtValueB06.Text = ""

                    'Display c1dateA06
                    'c1dateA06.Visible = True
                    c1dateA06.Top = c1dateA05.Top + LM_H
                    c1dateA06.Left = c1dateA05.Left
                    c1dateA06.Value = Now

                    'Display c1dateB06
                    'c1dateB06.Visible = True
                    c1dateB06.Top = c1dateB05.Top + LM_H
                    c1dateB06.Left = c1dateB05.Left
                    c1dateB06.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk06.Top = chk05.Top + LM_H
                    chk06.Left = chk01.Left
                    chk06.Checked = False


                    CboField06.Focus()
                Case 7

                    'Display Label7
                    Label7.Visible = True
                    Label7.Top = Label6.Top + LM_H
                    Label7.Left = Label6.Left
                    Label7.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label7.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label7.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label7.Visible = True

                    'Display cboField7
                    CboField07.Visible = True
                    CboField07.Top = CboField06.Top + LM_H
                    CboField07.Left = CboField06.Left
                    CboField07.Visible = True

                    'Display cboOperator7
                    CboOperator07.Visible = True
                    CboOperator07.Top = CboOperator06.Top + LM_H
                    CboOperator07.Left = CboOperator06.Left
                    CboOperator07.Enabled = False

                    'Display txtValueA7
                    TxtValueA07.Visible = True
                    TxtValueA07.Enabled = False
                    TxtValueA07.Top = TxtValueA06.Top + LM_H
                    TxtValueA07.Left = TxtValueA06.Left
                    TxtValueA07.Text = ""

                    'Display txtValueB7
                    'TxtValueB07.Visible = True
                    TxtValueB07.Top = TxtValueB06.Top + LM_H
                    TxtValueB07.Left = TxtValueA01.Left
                    TxtValueB07.Text = ""

                    'Display c1dateA07
                    'c1dateA07.Visible = True
                    c1dateA07.Top = c1dateA06.Top + LM_H
                    c1dateA07.Left = c1dateA06.Left
                    c1dateA07.Value = Now

                    'Display c1dateB07
                    'c1dateB07.Visible = True
                    c1dateB07.Top = c1dateB06.Top + LM_H
                    c1dateB07.Left = c1dateB06.Left
                    c1dateB07.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk07.Top = chk06.Top + LM_H
                    chk07.Left = chk01.Left
                    chk07.Checked = False


                    CboField07.Focus()
                Case 8

                    'Display Label8
                    Label8.Visible = True
                    Label8.Top = Label7.Top + LM_H
                    Label8.Left = Label7.Left
                    Label8.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label8.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label8.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label8.Visible = True

                    'Display cboField8
                    CboField08.Visible = True
                    CboField08.Top = CboField07.Top + LM_H
                    CboField08.Left = CboField07.Left
                    CboField08.Visible = True

                    'Display cboOperator8
                    CboOperator08.Visible = True
                    CboOperator08.Top = CboOperator07.Top + LM_H
                    CboOperator08.Left = CboOperator07.Left
                    CboOperator08.Enabled = False

                    'Display txtValueA8
                    TxtValueA08.Visible = True
                    TxtValueA08.Enabled = False
                    TxtValueA08.Top = TxtValueA07.Top + LM_H
                    TxtValueA08.Left = TxtValueA07.Left
                    TxtValueA08.Text = ""

                    'Display txtValueB8
                    'TxtValueB08.Visible = True
                    TxtValueB08.Top = TxtValueB07.Top + LM_H
                    TxtValueB08.Left = TxtValueA01.Left
                    TxtValueB08.Text = ""

                    'Display c1dateA08
                    'c1dateA08.Visible = True
                    c1dateA08.Top = c1dateA07.Top + LM_H
                    c1dateA08.Left = c1dateA07.Left
                    c1dateA08.Value = Now

                    'Display c1dateB08
                    'c1dateB08.Visible = True
                    c1dateB08.Top = c1dateB07.Top + LM_H
                    c1dateB08.Left = c1dateB07.Left
                    c1dateB08.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk08.Top = chk07.Top + LM_H
                    chk08.Left = chk01.Left
                    chk08.Checked = False


                    CboField08.Focus()
                Case 9

                    'Display Label9
                    Label9.Visible = True
                    Label9.Top = Label8.Top + LM_H
                    Label9.Left = Label8.Left
                    Label9.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label9.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label9.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label9.Visible = True

                    'Display cboField9
                    CboField09.Visible = True
                    CboField09.Top = CboField08.Top + LM_H
                    CboField09.Left = CboField08.Left
                    CboField09.Visible = True

                    'Display cboOperator9
                    CboOperator09.Visible = True
                    CboOperator09.Top = CboOperator08.Top + LM_H
                    CboOperator09.Left = CboOperator08.Left
                    CboOperator09.Enabled = False

                    'Display txtValueA9
                    TxtValueA09.Visible = True
                    TxtValueA09.Enabled = False
                    TxtValueA09.Top = TxtValueA08.Top + LM_H
                    TxtValueA09.Left = TxtValueA08.Left
                    TxtValueA09.Text = ""

                    'Display txtValueB9
                    'TxtValueB09.Visible = True
                    TxtValueB09.Top = TxtValueB08.Top + LM_H
                    TxtValueB09.Left = TxtValueA01.Left
                    TxtValueB09.Text = ""

                    'Display c1dateA09
                    'c1dateA09.Visible = True
                    c1dateA09.Top = c1dateA08.Top + LM_H
                    c1dateA09.Left = c1dateA08.Left
                    c1dateA09.Value = Now

                    'Display c1dateB09
                    'c1dateB09.Visible = True
                    c1dateB09.Top = c1dateB08.Top + LM_H
                    c1dateB09.Left = c1dateB08.Left
                    c1dateB09.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk09.Top = chk08.Top + LM_H
                    chk09.Left = chk01.Left
                    chk09.Checked = False


                    CboField09.Focus()
                Case 10

                    'Display Label10
                    Label10.Visible = True
                    Label10.Top = Label9.Top + LM_H
                    Label10.Left = Label9.Left
                    Label10.Text = sOperator

                    Select Case sOperator
                        Case LM_AND
                            Label10.ForeColor = System.Drawing.Color.Red
                        Case LM_OR
                            Label10.ForeColor = System.Drawing.Color.Blue
                    End Select

                    Label10.Visible = True

                    'Display cboField10
                    CboField10.Visible = True
                    CboField10.Top = CboField09.Top + LM_H
                    CboField10.Left = CboField09.Left
                    CboField10.Visible = True

                    'Display cboOperator10
                    CboOperator10.Visible = True
                    CboOperator10.Top = CboOperator09.Top + LM_H
                    CboOperator10.Left = CboOperator09.Left
                    CboOperator10.Enabled = False

                    'Display txtValueA10
                    TxtValueA10.Visible = True
                    TxtValueA10.Enabled = False
                    TxtValueA10.Top = TxtValueA09.Top + LM_H
                    TxtValueA10.Left = TxtValueA09.Left
                    TxtValueA10.Text = ""

                    'Display txtValueB10
                    'TxtValueB10.Visible = True
                    TxtValueB10.Top = TxtValueB09.Top + LM_H
                    TxtValueB10.Left = TxtValueA01.Left
                    TxtValueB10.Text = ""

                    'Display c1dateA10
                    ''c1dateA10.Visible = True
                    c1dateA10.Top = c1dateA09.Top + LM_H
                    c1dateA10.Left = c1dateA09.Left
                    c1dateA10.Value = Now

                    'Display c1dateB10
                    'c1dateB10.Visible = True
                    c1dateB10.Top = c1dateB09.Top + LM_H
                    c1dateB10.Left = c1dateB09.Left
                    c1dateB10.Value = Now

                    'Update 02/07/2010: Display Checkbox
                    chk10.Top = chk09.Top + LM_H
                    chk10.Left = chk01.Left
                    chk10.Checked = False


                    CboField10.Focus()
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub

    Private Sub DeleteRow()
        Try
            'EnableButton(True)

            If idx = 2 Then
                ButBack.Enabled = False
                btnBackFirst.Enabled = False
            Else
                ButBack.Enabled = True
                btnBackFirst.Enabled = True
            End If

            Me.Height = Me.Height - LM_H
            GroupBox1.Height = GroupBox1.Height - LM_H
            ButAnd.Top = ButAnd.Top - LM_H
            ButOr.Top = ButOr.Top - LM_H
            ButBack.Top = ButBack.Top - LM_H
            ButFind.Top = ButFind.Top - LM_H
            'ButHelp.Top = ButHelp.Top - LM_H
            ButClose.Top = ButClose.Top - LM_H

            btnPrevious.Top = btnBackFirst.Top - LM_H
            btnBackFirst.Top = btnBackFirst.Top - LM_H
            Select Case idx
                Case 2
                    VisibleOneRow(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, Label2, chk02, False)
                    EnabledControl_CboFieldKeyDelete(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, chk02)
                Case 3
                    VisibleOneRow(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, Label3, chk03, False)
                    EnabledControl_CboFieldKeyDelete(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, chk03)
                Case 4
                    VisibleOneRow(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, Label4, chk04, False)
                    EnabledControl_CboFieldKeyDelete(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, chk04)
                Case 5
                    VisibleOneRow(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, Label5, chk05, False)
                    EnabledControl_CboFieldKeyDelete(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, chk05)
                Case 6
                    VisibleOneRow(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, Label6, chk06, False)
                    EnabledControl_CboFieldKeyDelete(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, chk06)
                Case 7
                    VisibleOneRow(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, Label7, chk07, False)
                    EnabledControl_CboFieldKeyDelete(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, chk07)
                Case 8
                    VisibleOneRow(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, Label8, chk08, False)
                    EnabledControl_CboFieldKeyDelete(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, chk08)
                Case 9
                    VisibleOneRow(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, Label9, chk09, False)
                    EnabledControl_CboFieldKeyDelete(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, chk09)
                Case 10
                    VisibleOneRow(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, Label10, chk10, False)
                    EnabledControl_CboFieldKeyDelete(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, chk10)
            End Select

            idx = idx - 1
            EnableButton(True)
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try

    End Sub

    Private Sub AddItemCboOperator(ByVal cboOperator As C1Combo, ByVal DataType As FinderTypeEnum, ByRef DataTypeOld As FinderTypeEnum)
        With cboOperator
            If .Text <> "" Then
                If DataType = DataTypeOld Then
                    Exit Sub
                Else
                    DataTypeOld = DataType
                End If
            Else
                DataTypeOld = DataType
            End If

            bChangeText = True

            .ClearItems()
            If DataType = FinderTypeEnum.lmFinderString Then
                .AddItem(LM_OP_STARTWITH)
                .AddItem(LM_OP_CONTAIN)
                .AddItem(LM_OP_NOTCONTAIN)
                .AddItem(LM_OP_ENDWITH)
                .AddItem(LM_OP_EQUAL)
                .AddItem(LM_OP_NOTEQUAL)
                .MaxDropDownItems = 7

                .Text = LM_OP_CONTAIN

                'Update 04/06/2010: nếu cột là TinyInt thì chỉ có =
            ElseIf DataType = FinderTypeEnum.lmFinderTinyInt Then
                .AddItem(LM_OP_EQUAL)
                .MaxDropDownItems = 7

                .Text = LM_OP_EQUAL

            Else
                .AddItem(LM_OP_EQUAL)
                .AddItem(LM_OP_GREATEREQUAL)
                .AddItem(LM_OP_GREATER)
                .AddItem(LM_OP_LESSEQUAL)
                .AddItem(LM_OP_LESS)
                .AddItem(LM_OP_BETWEEN)
                .AddItem(LM_OP_NOTBETWEEN)
                .AddItem(LM_OP_NOTEQUAL)
                .MaxDropDownItems = 10

                If DataType = FinderTypeEnum.lmFinderDate Then
                    .Text = LM_OP_BETWEEN
                Else
                    .Text = LM_OP_EQUAL
                End If

            End If

            .ColumnHeaders = False
            .DropDownWidth = 132 '140
            .ExtendRightColumn = True
            .AutoDropDown = True
            .AutoCompletion = True
            .LimitToList = True
            .DropdownPosition = DropdownPositionEnum.LeftDown
            .Splits(0).DisplayColumns(0).Width = 100

            .Enabled = True
            .Focus()
        End With
    End Sub

    Private Sub StuffOperatorCombo(ByVal cboTemp As C1Combo, ByVal DataType As FinderTypeEnum)
        Try
            Select Case cboTemp.Name
                Case "CboField01"
                    AddItemCboOperator(CboOperator01, DataType, DataTypeField01)
                Case "CboField02"
                    AddItemCboOperator(CboOperator02, DataType, DataTypeField02)
                Case "CboField03"
                    AddItemCboOperator(CboOperator03, DataType, DataTypeField03)
                Case "CboField04"
                    AddItemCboOperator(CboOperator04, DataType, DataTypeField04)
                Case "CboField05"
                    AddItemCboOperator(CboOperator05, DataType, DataTypeField05)
                Case "CboField06"
                    AddItemCboOperator(CboOperator06, DataType, DataTypeField06)
                Case "CboField07"
                    AddItemCboOperator(CboOperator07, DataType, DataTypeField07)
                Case "CboField08"
                    AddItemCboOperator(CboOperator08, DataType, DataTypeField08)
                Case "CboField09"
                    AddItemCboOperator(CboOperator09, DataType, DataTypeField09)
                Case "CboField10"
                    AddItemCboOperator(CboOperator10, DataType, DataTypeField10)
            End Select

        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try

    End Sub

    Private Sub FillDataToArray()
        Dim i, intRowNum As Integer
        Dim arr() As Object

        intRowNum = UBound(arrAdvanced)

        ReDim xa(intRowNum)

        For i = 0 To intRowNum

            arr = Microsoft.VisualBasic.Split(arrAdvanced(i).ToString, LM_DLM)
            xa(i) = New XArray
            xa(i).Col1 = arr(0)
            xa(i).Col2 = CType(arr(1), Object)
            xa(i).Col3 = CType(arr(2), Object)
            xa(i).Col4 = CType(arr(3), Object)
            xa(i).Col5 = CType(arr(4), Object)

        Next i

        'BEGIN 
        Dim dc As DataColumn
        Dim dr As DataRow
        dtXA1 = New DataTable("xa")
        dc = New DataColumn("Col1", Type.GetType("System.String"))
        dtXA1.Columns.Add(dc)
        dc = New DataColumn("Col2", Type.GetType("System.String"))
        dtXA1.Columns.Add(dc)
        dc = New DataColumn("Col3", Type.GetType("System.String"))
        dtXA1.Columns.Add(dc)
        dc = New DataColumn("Col4", Type.GetType("System.String"))
        dtXA1.Columns.Add(dc)
        dc = New DataColumn("Col5", Type.GetType("System.String"))
        dtXA1.Columns.Add(dc)
        For i = 0 To intRowNum
            dr = dtXA1.NewRow()
            'dr(0) = xa(i).Col1
            If _useUnicode AndAlso _findServer Then ' Nếu là Unicode và tìm kiếm theo Server thì Convnert dữ liệu sang Unicode
                dr(0) = ConvertVniToUnicode(xa(i).Col1.ToString)
            Else
                dr(0) = xa(i).Col1
            End If

            dr(1) = xa(i).Col2
            dr(2) = xa(i).Col3
            dr(3) = xa(i).Col4
            dr(4) = xa(i).Col5
            dtXA1.Rows.Add(dr)
        Next
        dtXA2 = dtXA1.Copy
        dtXA3 = dtXA1.Copy
        dtXA4 = dtXA1.Copy
        dtXA5 = dtXA1.Copy
        dtXA6 = dtXA1.Copy
        dtXA7 = dtXA1.Copy
        dtXA8 = dtXA1.Copy
        dtXA9 = dtXA1.Copy
        dtXA10 = dtXA1.Copy
        'AND 
    End Sub

    Dim dtXA1 As DataTable = Nothing
    Dim dtXA2 As DataTable = Nothing
    Dim dtXA3 As DataTable = Nothing
    Dim dtXA4 As DataTable = Nothing
    Dim dtXA5 As DataTable = Nothing
    Dim dtXA6 As DataTable = Nothing
    Dim dtXA7 As DataTable = Nothing
    Dim dtXA8 As DataTable = Nothing
    Dim dtXA9 As DataTable = Nothing
    Dim dtXA10 As DataTable = Nothing

    Private Sub LoadCboField(ByVal CboField As C1.Win.C1List.C1Combo, ByVal dtXA As DataTable)

        With CboField
            'LoadDataSource(CboField, dtXA)
            .DataSource = dtXA

            .DropDownWidth = 200
            .ExtendRightColumn = True
            .AutoDropDown = True
            .AutoCompletion = True
            .LimitToList = False 'True Update 08/07/2010
            .MaxDropDownItems = 10 ' CType(intRowNum + 2, Short)' Date: 05/12/2007
            .ColumnHeaders = False
            .Splits(0).DisplayColumns(0).Width = 100
            .Splits(0).DisplayColumns(1).Visible = False
            .Splits(0).DisplayColumns(2).Visible = False
            .Splits(0).DisplayColumns(3).Visible = False
            .Splits(0).DisplayColumns(4).Visible = False
            .DropdownPosition = DropdownPositionEnum.LeftDown

            .DisplayMember = dtXA.Columns(0).ColumnName
            .ValueMember = dtXA.Columns(1).ColumnName

            If _useUnicode Then
                .Font = FontUnicode()
                .EditorFont = FontUnicode()
            End If
        End With

    End Sub

    Private Sub StuffFieldNameCombo()
        Try

            Dim intRowNum As Integer

            intRowNum = UBound(arrAdvanced)

            FillDataToArray()

            LoadCboField(CboField01, dtXA1)
            LoadCboField(CboField02, dtXA2)
            LoadCboField(CboField03, dtXA3)
            LoadCboField(CboField04, dtXA4)
            LoadCboField(CboField05, dtXA5)
            LoadCboField(CboField06, dtXA6)
            LoadCboField(CboField07, dtXA7)
            LoadCboField(CboField08, dtXA8)
            LoadCboField(CboField09, dtXA9)
            LoadCboField(CboField10, dtXA10)

        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub D99F0006_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                Application.DoEvents()
                If Me.ActiveControl.Name <> "" Then
                    If CInt(Me.ActiveControl.Name.Substring(Me.ActiveControl.Name.Length - 2)) = idx Then
                        If bFlagBETWEEN Then ' Tìm kiếm trong hay ngoài khoảng
                            If Me.ActiveControl.Name.Substring(Me.ActiveControl.Name.Length - 3, 1) = "B" Then
                                'Hoàng Long: 24/06/2009
                                If Not CheckTxtValue(idx) Then Exit Sub
                                ButFind_Click(sender, e)
                            Else
                                UseEnterAsTab(Me)
                            End If
                        Else
                            'Hoàng Long: 24/06/2009
                            If Not CheckTxtValue(idx) Then Exit Sub
                            ButFind_Click(sender, e)
                        End If
                    Else
                        UseEnterAsTab(Me)
                    End If
                Else
                    UseEnterAsTab(Me)
                End If
                Application.DoEvents()
            Case Keys.F8
                Application.DoEvents()
                btnPrevious_Click(Nothing, Nothing)
                Application.DoEvents()
            Case Keys.F9
                Application.DoEvents()
                If ButAnd.Enabled Then
                    ButAnd_Click(Nothing, Nothing)
                End If
                Application.DoEvents()
            Case Keys.F10
                Application.DoEvents()
                If ButOr.Enabled Then
                    ButOr_Click(Nothing, Nothing)
                End If
                Application.DoEvents()
            Case Keys.F11
                Application.DoEvents()
                If ButBack.Enabled Then
                    ButBack_Click(Nothing, Nothing)
                End If
                Application.DoEvents()
            Case Keys.F12
                Application.DoEvents()
                If btnBackFirst.Enabled Then
                    btnBackFirst_Click(Nothing, Nothing)
                End If
                Application.DoEvents()
        End Select

        'If e.Alt Then
        '    Application.DoEvents()
        '    e.SuppressKeyPress = True
        '    Application.DoEvents()
        'End If
    End Sub

    Private Function CheckKeyPressString(ByVal KeyChar As Char) As Boolean
        'Chuỗi không cho gõ ký tự đặc biệt này
        If KeyChar = Chr(91) Then Return True ' [
        If KeyChar = Chr(93) Then Return True ']

    End Function

    Private Sub TxtValueA1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA01.KeyPress

        Select Case CType(CboField01.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                'Update 04/06/2010
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
                'e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Custom, "01")
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select

    End Sub

    Private Sub TxtValueA2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA02.KeyPress

        Select Case CType(CboField02.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select

    End Sub

    Private Sub TxtValueA3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA03.KeyPress
        Select Case CType(CboField03.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueA4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA04.KeyPress

        Select Case CType(CboField04.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueA5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA05.KeyPress
        Select Case CType(CboField05.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueA6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA06.KeyPress
        Select Case CType(CboField06.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueA7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA07.KeyPress
        Select Case CType(CboField07.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueA8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA08.KeyPress
        Select Case CType(CboField08.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueA9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA09.KeyPress
        Select Case CType(CboField09.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueA10_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueA10.KeyPress
        Select Case CType(CboField10.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString
                e.Handled = CheckKeyPressString(e.KeyChar)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueB1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB01.KeyPress
        Select Case CType(CboField01.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select

    End Sub

    Private Sub TxtValueB2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB02.KeyPress
        Select Case CType(CboField02.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueB3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB03.KeyPress
        Select Case CType(CboField03.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select

    End Sub

    Private Sub TxtValueB4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB04.KeyPress
        Select Case CType(CboField04.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueB5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB05.KeyPress
        Select Case CType(CboField05.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueB6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB06.KeyPress
        Select Case CType(CboField06.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueB7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB07.KeyPress
        Select Case CType(CboField07.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueB8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB08.KeyPress
        Select Case CType(CboField08.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueB9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB09.KeyPress
        Select Case CType(CboField09.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    Private Sub TxtValueB10_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValueB10.KeyPress
        Select Case CType(CboField10.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderTinyInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case FinderTypeEnum.lmFinderInt
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberSign)
            Case FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
        End Select
    End Sub

    'Description: Đổi các sự kiện LostFocus của TextBox thành Validated.
    '------------ Sửa hàm LostFocusTxtValue thành Function kiểu Bool, bỏ set về trắng khi dữ liệu không hợp lệ
#Region "Hoàng Long, 24/06/2009"

    Private Function LostFocusTxtValue(ByVal TxtValue As TextBox, ByVal CboField As C1.Win.C1List.C1Combo) As Boolean
        If ButClose.Focused Then
            Exit Function
        End If

        Select Case CType(CboField.Columns(2).Value, D99D0041.FinderTypeEnum)
            Case FinderTypeEnum.lmFinderString

                If TxtValue.Text.Trim <> "" AndAlso TxtValue.Text.Trim.Substring(0, 1) = ";" Then
                    TxtValue.Focus()
                    D99C0008.MsgL3(r("Gia_tri_khong_hop_le"))
                    Return False
                End If

            Case FinderTypeEnum.lmFinderTinyInt 'Minh Hòa update 14/09/2010: là checkbox trên lưới không kiểm tra

                'If Trim(TxtValue.Text) <> "" Then
                '    If Not IsNumeric(TxtValue.Text) OrElse TxtValue.Text.Contains(".") Then
                '        TxtValue.Focus()
                '        D99C0008.MsgL3(sMsgNhapSoChuaDung)
                '        Return False
                '    End If

                '    If Val(TxtValue.Text) > MaxTinyInt Then
                '        TxtValue.Focus()
                '        D99C0008.MsgL3(sMsgNhapSoQuaLon)
                '        Return False
                '    End If
                'End If

            Case FinderTypeEnum.lmFinderInt

                If Trim(TxtValue.Text) <> "" Then
                    If Not IsNumeric(TxtValue.Text) OrElse TxtValue.Text.Contains(".") Then
                        TxtValue.Focus()
                        D99C0008.MsgL3(sMsgNhapSoChuaDung)
                        Return False
                    End If

                    If Val(TxtValue.Text) > MaxInt Then
                        TxtValue.Focus()
                        D99C0008.MsgL3(sMsgNhapSoQuaLon)
                        Return False
                    End If
                End If


            Case FinderTypeEnum.lmFinderSmallMoney

                If Trim(TxtValue.Text) <> "" Then
                    If Not IsNumeric(TxtValue.Text) Then
                        TxtValue.Focus()
                        D99C0008.MsgL3(sMsgNhapSoChuaDung)
                        Return False
                    End If

                    If TxtValue.Text.Contains(".") Then ' Có 4 số lẻ
                        If TxtValue.Text.Substring(0, TxtValue.Text.LastIndexOf(".")).Length > MaxSmallMoney.ToString.Length - 1 OrElse TxtValue.Text.Substring(TxtValue.Text.LastIndexOf(".") + 1).Length > 4 Then
                            TxtValue.Focus()
                            D99C0008.MsgL3(sMsgNhapSoChuaDung)
                            Return False
                        End If
                    Else
                        If TxtValue.Text.Length > MaxSmallMoney.ToString.Length Then
                            TxtValue.Focus()
                            D99C0008.MsgL3(sMsgNhapSoQuaLon)
                            Return False
                        End If
                    End If

                    If Val(TxtValue.Text) > MaxSmallMoney Then
                        TxtValue.Focus()
                        D99C0008.MsgL3(sMsgNhapSoQuaLon)
                        Return False
                    End If
                End If

            Case FinderTypeEnum.lmFinderMoney

                If Trim(TxtValue.Text) <> "" Then
                    If Not IsNumeric(TxtValue.Text) Then
                        TxtValue.Focus()
                        D99C0008.MsgL3(sMsgNhapSoChuaDung)
                        Return False
                    End If

                    If TxtValue.Text.Contains(".") Then ' Có 4 số lẻ
                        If TxtValue.Text.Substring(0, TxtValue.Text.LastIndexOf(".")).Length > MaxMoney.ToString.Length - 1 OrElse TxtValue.Text.Substring(TxtValue.Text.LastIndexOf(".") + 1).Length > 4 Then
                            TxtValue.Focus()
                            D99C0008.MsgL3(sMsgNhapSoChuaDung)
                            Return False
                        End If
                    Else
                        If TxtValue.Text.Length > MaxMoney.ToString.Length Then
                            TxtValue.Focus()
                            D99C0008.MsgL3(sMsgNhapSoQuaLon)
                            Return False
                        End If
                    End If

                    If Val(TxtValue.Text) > MaxMoney Then
                        TxtValue.Focus()
                        D99C0008.MsgL3(sMsgNhapSoQuaLon)
                        Return False
                    End If
                End If

            Case FinderTypeEnum.lmFinderNumber 'Minh Hòa update 15/07/2009: Tương đương với Decimal trong SQL

                If Trim(TxtValue.Text) <> "" Then
                    If Not IsNumeric(TxtValue.Text) Then
                        TxtValue.Focus()
                        D99C0008.MsgL3(sMsgNhapSoChuaDung)
                        Return False
                    End If

                    If TxtValue.Text.Contains(".") Then ' có 8 Số lẻ
                        If TxtValue.Text.Substring(0, TxtValue.Text.LastIndexOf(".")).Length > MaxDecimal.ToString.Length - 1 OrElse TxtValue.Text.Substring(TxtValue.Text.LastIndexOf(".") + 1).Length > 8 Then
                            TxtValue.Focus()
                            D99C0008.MsgL3(sMsgNhapSoChuaDung)
                            Return False
                        End If
                    Else
                        If TxtValue.Text.Length > MaxDecimal.ToString.Length - 1 Then
                            TxtValue.Focus()
                            D99C0008.MsgL3(sMsgNhapSoQuaLon)
                            Return False
                        End If
                    End If

                    If CDec(TxtValue.Text) > MaxDecimal Then
                        TxtValue.Focus()
                        D99C0008.MsgL3(sMsgNhapSoQuaLon)
                        Return False
                    End If
                End If
        End Select
        Return True
    End Function

    Private Sub TxtValueA1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA01.Validated
        LostFocusTxtValue(TxtValueA01, CboField01)
    End Sub

    Private Sub TxtValueA2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA02.Validated
        LostFocusTxtValue(TxtValueA02, CboField02)
    End Sub

    Private Sub TxtValueA3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA03.Validated
        LostFocusTxtValue(TxtValueA03, CboField03)
    End Sub

    Private Sub TxtValueA4_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA04.Validated
        LostFocusTxtValue(TxtValueA04, CboField04)
    End Sub

    Private Sub TxtValueA5_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA05.Validated
        LostFocusTxtValue(TxtValueA05, CboField05)
    End Sub

    Private Sub TxtValueA6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA06.Validated
        LostFocusTxtValue(TxtValueA06, CboField06)
    End Sub

    Private Sub TxtValueA7_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA07.Validated
        LostFocusTxtValue(TxtValueA07, CboField07)
    End Sub

    Private Sub TxtValueA8_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA08.Validated
        LostFocusTxtValue(TxtValueA08, CboField08)
    End Sub

    Private Sub TxtValueA9_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA09.Validated
        LostFocusTxtValue(TxtValueA09, CboField09)
    End Sub

    Private Sub TxtValueA10_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueA10.Validated
        LostFocusTxtValue(TxtValueA10, CboField10)
    End Sub

    Private Sub TxtValueB1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB01.Validated
        LostFocusTxtValue(TxtValueB01, CboField01)
    End Sub

    Private Sub TxtValueB2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB02.Validated
        LostFocusTxtValue(TxtValueB02, CboField02)
    End Sub

    Private Sub TxtValueB3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB03.Validated
        LostFocusTxtValue(TxtValueB03, CboField03)
    End Sub

    Private Sub TxtValueB4_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB04.Validated
        LostFocusTxtValue(TxtValueB04, CboField04)
    End Sub

    Private Sub TxtValueB5_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB05.Validated
        LostFocusTxtValue(TxtValueB05, CboField05)
    End Sub

    Private Sub TxtValueB6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB06.Validated
        LostFocusTxtValue(TxtValueB06, CboField06)
    End Sub

    Private Sub TxtValueB7_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB07.Validated
        LostFocusTxtValue(TxtValueB07, CboField07)
    End Sub

    Private Sub TxtValueB8_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB08.Validated
        LostFocusTxtValue(TxtValueB08, CboField08)
    End Sub

    Private Sub TxtValueB9_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB09.Validated
        LostFocusTxtValue(TxtValueB09, CboField09)
    End Sub

    Private Sub TxtValueB10_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtValueB10.Validated
        LostFocusTxtValue(TxtValueB10, CboField10)
    End Sub
#End Region

    Private Sub Center(ByVal aForm As Form)
        'Canh giữa màn hình
        Dim workingRectangle As System.Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea
        Dim tempPoint As New Point(CType((workingRectangle.Width - aForm.Width) / 2, Integer), CType((workingRectangle.Height - aForm.Height) / 2, Integer))

        aForm.DesktopLocation = tempPoint

    End Sub

    Private Function DateSave(ByVal [Date] As String) As String
        If [Date] = "" Then Return "NULL"
        Dim dDate As Date = CType([Date], Date)
        Return dDate.ToString("MM/dd/yyyy")
        'Return dDate.ToString("MM/dd/yyyy HH:mm:ss")
        '"MM/dd/yyyy HH:mm:ss"
    End Function

    Private Function DateSave(ByVal [Date] As Object) As String
        If IsDBNull([Date]) Then Return "NULL"
        Return DateSave([Date].ToString)
    End Function

    Private Sub EnabledControl_CboFieldKeyDelete(ByVal CboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal chk As System.Windows.Forms.CheckBox)
        CboField.Text = ""

        CboOperator.Enabled = False
        CboOperator.Text = ""
        TxtValueA.Enabled = False
        TxtValueA.Text = ""
        TxtValueB.Enabled = False
        TxtValueB.Text = ""
        c1dateA.Enabled = False
        c1dateA.Value = Now
        c1dateB.Enabled = False
        c1dateB.Value = Now
        chk.Enabled = False
        chk.Checked = False

        ButFind.Enabled = False
        'If idx = 1 Then ButFind.Enabled = False
    End Sub

    Private Sub EnabledControl_CboOperatorKeyDelete(ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit)
        CboOperator.Text = ""
        TxtValueA.Enabled = False
        TxtValueA.Text = ""
        TxtValueB.Enabled = False
        TxtValueB.Text = ""
        c1dateA.Enabled = False
        c1dateA.Value = Now
        c1dateB.Enabled = False
        c1dateB.Value = Now
        ButFind.Enabled = False
    End Sub



    Private Sub CboOperator01_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator01.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01)
        End If
    End Sub

    Private Sub CboOperator02_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator02.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02)
        End If
    End Sub

    Private Sub CboOperator03_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator03.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03)
        End If
    End Sub

    Private Sub CboOperator04_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator04.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04)
        End If
    End Sub

    Private Sub CboOperator05_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator05.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05)
        End If
    End Sub

    Private Sub CboOperator06_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator06.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06)
        End If
    End Sub

    Private Sub CboOperator07_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator07.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07)
        End If
    End Sub

    Private Sub CboOperator08_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator08.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08)
        End If
    End Sub

    Private Sub CboOperator09_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator09.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09)
        End If
    End Sub

    Private Sub CboOperator10_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CboOperator10.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            EnabledControl_CboOperatorKeyDelete(CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10)
        End If
    End Sub

    Private Sub D99F0006_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If TxtValueA01.Enabled Then TxtValueA01.Focus()
    End Sub


#Region "Hoang Long"
    Private Sub btnBackFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackFirst.Click
        DeleteAll()
        EnableButton(True)
        idx = 1
        Center(Me)
    End Sub

    Private Sub DeleteAll()
        Dim iHeight As Integer = LM_H * (idx - 1)
        ButOr.Top -= iHeight
        ButAnd.Top -= iHeight
        ButFind.Top -= iHeight
        ButClose.Top -= iHeight
        ButBack.Top -= iHeight
        btnPrevious.Top -= iHeight
        btnBackFirst.Top -= iHeight
        GroupBox1.Height -= iHeight
        Me.Height -= iHeight
        VisibleAll()
        btnBackFirst.Enabled = False
        ButBack.Enabled = False
        CboField01.Focus()
    End Sub

    Private Sub VisibleOneRow(ByVal cboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal Label As Label, ByVal chk As System.Windows.Forms.CheckBox, ByVal Flag As Boolean)
        cboField.Visible = Flag
        CboOperator.Visible = Flag
        TxtValueA.Visible = Flag
        TxtValueB.Visible = Flag
        c1dateA.Visible = Flag
        c1dateB.Visible = Flag
        Label.Visible = Flag
        chk.Visible = Flag
    End Sub

    Private Sub VisibleAll()

        VisibleOneRow(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, Label2, chk02, False)
        EnabledControl_CboFieldKeyDelete(CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, chk02)

        VisibleOneRow(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, Label3, chk03, False)
        EnabledControl_CboFieldKeyDelete(CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, chk03)

        VisibleOneRow(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, Label4, chk04, False)
        EnabledControl_CboFieldKeyDelete(CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, chk04)

        VisibleOneRow(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, Label5, chk05, False)
        EnabledControl_CboFieldKeyDelete(CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, chk05)

        VisibleOneRow(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, Label6, chk06, False)
        EnabledControl_CboFieldKeyDelete(CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, chk06)

        VisibleOneRow(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, Label7, chk07, False)
        EnabledControl_CboFieldKeyDelete(CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, chk07)

        VisibleOneRow(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, Label8, chk08, False)
        EnabledControl_CboFieldKeyDelete(CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, chk08)

        VisibleOneRow(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, Label9, chk09, False)
        EnabledControl_CboFieldKeyDelete(CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, chk09)

        VisibleOneRow(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, Label10, chk10, False)
        EnabledControl_CboFieldKeyDelete(CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, chk10)

    End Sub

    Private Sub LoadTemplateID()
        Dim sWhereFieldName As String
        If _useUnicode Then
            sWhereFieldName = " And FieldNameU <> ''"
        Else
            sWhereFieldName = " And FieldName <> ''"
        End If

        Dim sSQL As String
        sSQL = "Select * from D91T1100 WITH(NOLOCK) " _
        & "Where FormID = " & SQLString(_formID) & " And Mode = " & SQLString(_mode) _
        & " And SubMode = '' And UserID = " & SQLString(gsUserID) _
        & sWhereFieldName _
        & " Order By CreateDate Desc, FieldOrderNo"

        Dim dt1 As DataTable = ReturnDataTable(sSQL)
        If dt1.Rows.Count > 0 Then
            'Lấy TemplateID của lần tìm kiếm gần nhất 
            iLastTemplateID = CInt(dt1.Rows(0).Item("TemplateID"))
            iTemplateID = iLastTemplateID
            Dim dt2 As DataTable = dt1.Copy
            dt2.DefaultView.RowFilter = "TemplateID = " & SQLString(iTemplateID)
            dt2.DefaultView.Sort = "FieldOrderNo"
            LoadLastFinder(dt2.DefaultView.ToTable)

            'Lấy TemplateID lớn nhất
            dt2 = dt1.Copy
            dt2.DefaultView.Sort = "TemplateID Desc"
            iMaxTemplateID = CInt(dt2.DefaultView.Item(0).Item("TemplateID").ToString)

            Center(Me)
            CboField01.Focus()
        End If
    End Sub

    Private Sub LoadLastFinder(Optional ByVal dt1 As DataTable = Nothing)
        'Load dữ liệu của lần tìm kiếm gần nhất        

        If dt1 Is Nothing Then
            Dim sSQL As String = "Select * from D91T1100 WITH(NOLOCK) Where TemplateID = " & SQLString(iTemplateID.ToString) _
            & " And FormID=" & SQLString(_formID) & " And Mode = " & SQLString(_mode) _
            & " And SubMode = '' And UserID = " & SQLString(gsUserID) _
            & " Order By FieldOrderNo"
            dt1 = ReturnDataTable(sSQL)
        End If

        '*** Update 20/05/2010: nếu mẫu thiết lập không nằm trong danh sách combo chứa các field tìm kiếm thì remove đi
        Dim bFlag As Boolean = False
        For j As Integer = dt1.Rows.Count - 1 To 0 Step -1
            bFlag = False
            For i As Integer = 0 To dtXA1.Rows.Count - 1
                'Update 02/08/2010: Unicode thì lấy FieldNameU, còn VNI thì lấy FieldName
                'If dt1.Rows(j).Item("FieldName").ToString = dtXA1.Rows(i).Item("Col2").ToString Then
                If dt1.Rows(j).Item("FieldName" & UnicodeJoin(_useUnicode)).ToString = dtXA1.Rows(i).Item("Col2").ToString Then
                    bFlag = True
                    Exit For
                End If
            Next
            If Not bFlag Then
                dt1.Rows.RemoveAt(j)
            End If
        Next
        '***

        For i As Integer = 0 To dt1.Rows.Count - 1
            idx = i
            Select Case i
                Case 0
                    LoadRow(dt1.Rows(i), CboField01, CboOperator01, TxtValueA01, TxtValueB01, c1dateA01, c1dateB01, Label1, chk01)
                Case 1
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField02, CboOperator02, TxtValueA02, TxtValueB02, c1dateA02, c1dateB02, Label2, chk02)
                Case 2
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField03, CboOperator03, TxtValueA03, TxtValueB03, c1dateA03, c1dateB03, Label3, chk03)
                Case 3
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField04, CboOperator04, TxtValueA04, TxtValueB04, c1dateA04, c1dateB04, Label4, chk04)
                Case 4
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField05, CboOperator05, TxtValueA05, TxtValueB05, c1dateA05, c1dateB05, Label5, chk05)
                Case 5
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField06, CboOperator06, TxtValueA06, TxtValueB06, c1dateA06, c1dateB06, Label6, chk06)
                Case 6
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField07, CboOperator07, TxtValueA07, TxtValueB07, c1dateA07, c1dateB07, Label7, chk07)
                Case 7
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField08, CboOperator08, TxtValueA08, TxtValueB08, c1dateA08, c1dateB08, Label8, chk08)
                Case 8
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField09, CboOperator09, TxtValueA09, TxtValueB09, c1dateA09, c1dateB09, Label9, chk09)
                Case 9
                    AddRow(dt1.Rows(i).Item("Operator").ToString)
                    LoadRow(dt1.Rows(i), CboField10, CboOperator10, TxtValueA10, TxtValueB10, c1dateA10, c1dateB10, Label10, chk10)
            End Select
        Next
        If idx = 0 Then
            idx = 1
        End If
    End Sub

    Private Sub LoadRow(ByVal iRow As DataRow, ByVal cboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal Label As Label, ByVal chk As System.Windows.Forms.CheckBox)
        With iRow

            cboField.SelectedValue = .Item("FieldName" & UnicodeJoin(_useUnicode)).ToString
            If cboField.SelectedValue Is Nothing OrElse cboField.SelectedValue.ToString = "" Then
                CboOperator.Text = ""
                TxtValueA.Text = ""
                TxtValueB.Text = ""
                c1dateA.Value = Nothing
                c1dateB.Value = Nothing
                Exit Sub
            End If
            CloseCboField(cboField, CboOperator, TxtValueA, TxtValueB, c1dateA, c1dateB, chk)

            Dim findex As Integer
            findex = CType(cboField.Columns(2).Text, Integer)

            'Update 12/07/2010
            If CInt(.Item("StrCompare").ToString) >= CboOperator.ListCount Then
                CboOperator.SelectedIndex = 0
            Else
                CboOperator.SelectedIndex = CInt(.Item("StrCompare").ToString)
            End If

            bChangeText = False
            CloseCboOperator(cboField, CboOperator, TxtValueA, TxtValueB, c1dateA, c1dateB, chk)
            Select Case findex
                Case FinderTypeEnum.lmFinderString
                    If _useUnicode Then
                        TxtValueA.Text = .Item("StrValueU").ToString
                    Else
                        TxtValueA.Text = .Item("StrValue").ToString
                    End If
                Case FinderTypeEnum.lmFinderDate
                    Select Case CboOperator.Text
                        Case LM_OP_BETWEEN, LM_OP_NOTBETWEEN
                            c1dateA.Value = SQLDateShow(.Item("DateFrom").ToString)
                            c1dateB.Value = SQLDateShow(.Item("DateTo").ToString)
                        Case LM_OP_EQUAL, LM_OP_GREATEREQUAL, LM_OP_GREATER, LM_OP_LESSEQUAL, LM_OP_LESS, LM_OP_NOTEQUAL
                            c1dateA.Value = SQLDateShow(.Item("DateFrom").ToString)
                    End Select

                Case FinderTypeEnum.lmFinderTinyInt ' Checkbox
                    chk.Checked = L3Bool(.Item("NumFrom").ToString)
                Case FinderTypeEnum.lmFinderInt, FinderTypeEnum.lmFinderTinyInt, FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney, FinderTypeEnum.lmFinderNumber
                    Select Case CboOperator.Text
                        Case LM_OP_BETWEEN, LM_OP_NOTBETWEEN
                            TxtValueA.Text = .Item("NumFrom").ToString
                            TxtValueB.Text = .Item("NumTo").ToString
                        Case LM_OP_EQUAL, LM_OP_GREATEREQUAL, LM_OP_GREATER, LM_OP_LESSEQUAL, LM_OP_LESS, LM_OP_NOTEQUAL
                            TxtValueA.Text = .Item("NumFrom").ToString
                    End Select
            End Select

            If .Item("Operator").ToString = "AND" Then
                Label.ForeColor = System.Drawing.Color.Red
                Label.Text = LM_AND
            Else
                Label.ForeColor = System.Drawing.Color.Blue
                Label.Text = LM_OR
            End If
        End With
    End Sub

    Private Function ReplaceSQLInsert(ByVal s As String, ByVal sFieldOrderNo As String, ByVal cboField As C1.Win.C1List.C1Combo, ByVal CboOperator As C1.Win.C1List.C1Combo, ByVal TxtValueA As System.Windows.Forms.TextBox, ByVal TxtValueB As System.Windows.Forms.TextBox, ByVal c1dateA As C1.Win.C1Input.C1DateEdit, ByVal c1dateB As C1.Win.C1Input.C1DateEdit, ByVal chk As System.Windows.Forms.CheckBox, Optional ByVal sOperator As String = "") As String
        Dim findex As Integer
        findex = CType(cboField.Columns(2).Text, Integer)
        s = s.Replace("{FieldOrderNo}", sFieldOrderNo)
        's = s.Replace("{FieldName}", SQLString(cboField.SelectedValue.ToString))
        If _useUnicode Then
            s = s.Replace("{FieldName}", SQLString(""))
            s = s.Replace("{FieldNameU}", SQLString(cboField.SelectedValue.ToString))
        Else
            s = s.Replace("{FieldName}", SQLString(cboField.SelectedValue.ToString))
            s = s.Replace("{FieldNameU}", SQLString(""))
        End If
        s = s.Replace("{StrCompare}", CboOperator.SelectedIndex.ToString)

        Select Case findex
            Case FinderTypeEnum.lmFinderString
                If TxtValueA.Text = " " Then ' Tìm khoảng trắng
                    s = s.Replace("{StrValue}", "' '")
                    s = s.Replace("{StrValueU}", "' '")
                Else
                    's = s.Replace("{StrValue}", SQLStringTextbox(TxtValueA, False))
                    's = s.Replace("{StrValueU}", SQLStringTextbox(TxtValueA, True))
                    If _useUnicode Then
                        s = s.Replace("{StrValue}", SQLString(""))
                        s = s.Replace("{StrValueU}", "N" & SQLString(TxtValueA.Text))
                    Else
                        s = s.Replace("{StrValue}", SQLString(TxtValueA.Text))
                        s = s.Replace("{StrValueU}", SQLString(""))
                    End If
                End If
                s = s.Replace("{DateFrom}", "''")
                s = s.Replace("{DateTo}", "''")
                s = s.Replace("{NumFrom}", "0")
                s = s.Replace("{NumTo}", "0")
            Case FinderTypeEnum.lmFinderDate
                Select Case CboOperator.Text
                    Case LM_OP_BETWEEN, LM_OP_NOTBETWEEN
                        s = s.Replace("{DateFrom}", SQLDateSave(c1dateA.Text))
                        s = s.Replace("{DateTo}", SQLDateSave(c1dateB.Text))
                    Case LM_OP_EQUAL, LM_OP_GREATEREQUAL, LM_OP_GREATER, LM_OP_LESSEQUAL, LM_OP_LESS, LM_OP_NOTEQUAL
                        s = s.Replace("{DateFrom}", SQLDateSave(c1dateA.Text))
                        s = s.Replace("{DateTo}", "''")
                End Select
                s = s.Replace("{StrValue}", "''")
                s = s.Replace("{StrValueU}", "''")
                s = s.Replace("{NumFrom}", "0")
                s = s.Replace("{NumTo}", "0")

            Case FinderTypeEnum.lmFinderTinyInt 'Minh Hòa update 02/07/2010: checkbox
                s = s.Replace("{NumFrom}", SQLNumber(chk.Checked))
                s = s.Replace("{NumTo}", "0")


                s = s.Replace("{StrValue}", "''")
                s = s.Replace("{StrValueU}", "''")
                s = s.Replace("{DateFrom}", "''")
                s = s.Replace("{DateTo}", "''")

            Case FinderTypeEnum.lmFinderInt, FinderTypeEnum.lmFinderMoney, FinderTypeEnum.lmFinderSmallMoney ', FinderTypeEnum.lmFinderNumber
                Select Case CboOperator.Text
                    Case LM_OP_BETWEEN, LM_OP_NOTBETWEEN
                        s = s.Replace("{NumFrom}", SQLNumber(TxtValueA.Text))
                        s = s.Replace("{NumTo}", SQLNumber(TxtValueB.Text))
                    Case LM_OP_EQUAL, LM_OP_GREATEREQUAL, LM_OP_GREATER, LM_OP_LESSEQUAL, LM_OP_LESS, LM_OP_NOTEQUAL
                        s = s.Replace("{NumFrom}", SQLNumber(TxtValueA.Text))
                        s = s.Replace("{NumTo}", "0")
                End Select
                s = s.Replace("{StrValue}", "''")
                s = s.Replace("{StrValueU}", "''")
                s = s.Replace("{DateFrom}", "''")
                s = s.Replace("{DateTo}", "''")

            Case FinderTypeEnum.lmFinderNumber 'Minh Hòa update 15/07/2009: Tương đương với Decimal trong SQL
                Select Case CboOperator.Text
                    Case LM_OP_BETWEEN, LM_OP_NOTBETWEEN
                        s = s.Replace("{NumFrom}", SQLDecimal(TxtValueA.Text))
                        s = s.Replace("{NumTo}", SQLDecimal(TxtValueB.Text))
                    Case LM_OP_EQUAL, LM_OP_GREATEREQUAL, LM_OP_GREATER, LM_OP_LESSEQUAL, LM_OP_LESS, LM_OP_NOTEQUAL
                        's = s.Replace("{NumFrom}", SQLNumber(TxtValueA.Text))
                        s = s.Replace("{NumFrom}", SQLDecimal(TxtValueA.Text))
                        s = s.Replace("{NumTo}", "0")
                End Select
                s = s.Replace("{StrValue}", "''")
                s = s.Replace("{StrValueU}", "''")
                s = s.Replace("{DateFrom}", "''")
                s = s.Replace("{DateTo}", "''")

        End Select
        If sOperator = LM_AND Then
            s = s.Replace("{Operator}", SQLString("AND"))
        ElseIf sOperator = LM_OR Then
            s = s.Replace("{Operator}", SQLString("OR"))
        Else
            s = s.Replace("{Operator}", SQLString(sOperator))
        End If
        Return s
    End Function

    Private Sub btnPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious.Click
        If iMaxTemplateID < 2 Then Exit Sub
        iTemplateID -= 1
        If iTemplateID = 0 Then
            iTemplateID = iMaxTemplateID
        End If
        btnBackFirst_Click(Nothing, Nothing)
        LoadLastFinder()
        Center(Me)
        CboField01.Focus()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD91T1100
    '# Created User: Nguyễn Hoàng Long
    '# Created Date: 26/02/2009 11:29:16
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD91T1100() As String
        Dim sSQL As String
        sSQL = "Delete From D91T1100 " _
        & "Where Mode = " & SQLString(_mode) _
        & " And SubMode = '' And UserID = " & SQLString(gsUserID) _
        & " And FormID = " & SQLString(_formID) _
        & " And TemplateId = " & SQLString(iTemplateID)
        Return sSQL
    End Function
#End Region

    'Hoàng Long: 24/06/2009
    'Description: Kiểm tra dữ liệu khi bấm Enter tại các TextBox trước khi tìm kiếm
    Private Function CheckTxtValue(ByVal index As Integer) As Boolean
        Select Case index
            Case 1
                Return LostFocusTxtValue(TxtValueA01, CboField01)
            Case 2
                Return LostFocusTxtValue(TxtValueA02, CboField02)
            Case 3
                Return LostFocusTxtValue(TxtValueA03, CboField03)
            Case 4
                Return LostFocusTxtValue(TxtValueA04, CboField04)
            Case 5
                Return LostFocusTxtValue(TxtValueA05, CboField05)
            Case 6
                Return LostFocusTxtValue(TxtValueA06, CboField06)
            Case 7
                Return LostFocusTxtValue(TxtValueA07, CboField07)
            Case 8
                Return LostFocusTxtValue(TxtValueA08, CboField08)
            Case 9
                Return LostFocusTxtValue(TxtValueA09, CboField09)
            Case 10
                Return LostFocusTxtValue(TxtValueA10, CboField10)
        End Select
        Return True
    End Function

End Class
