Imports System.Drawing.Printing

Public Class frmBegin
    Inherits System.Windows.Forms.Form
    Dim searchDS As DataSet = New DataSet
    Dim searchDA As SqlClient.SqlDataAdapter

    Dim AnyData() As String

    Dim chkBill As Boolean = False
    Dim ONo As String
    Dim BNo As String
    'Dim strDocCon As String = ""
    Dim intChk As Integer = 0
    Dim BNo2 As String = ""
    Dim StkID As String
    Dim StkName As String

    Dim StkPrice As Double
    Dim StkPrPerKg As Double
    Dim stkStock As Double = 0
    Dim BillDate As Date
    Dim cusCode As String = ""
    Dim CusName As String
    Dim Zone As String
    Dim BNum As Double
    Dim NumQ2 As Double = 0
    Dim BPrice As Double
    Dim BBill As Double

    Dim BSum As Double
    Dim lvi As ListViewItem
    Dim item_Name As New ListViewItem("Car_Name", 0)
    Dim cusAddr As String = ""
    Dim cusDetail As String = ""
    Dim cusDetail2 As String = ""
    Dim cusSales As String = ""

    Dim sumW As Double = 0


    Dim chkTB As Boolean
    Dim Transport As String
    Dim i_row As Integer
    Dim chkRow As Boolean = False


    Friend WithEvents lbStkName As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtQty2 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lbQty1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents chkCheck As System.Windows.Forms.CheckBox
    Friend WithEvents lbStkCode As System.Windows.Forms.Label
    Friend WithEvents lbArCode As System.Windows.Forms.Label
    Friend WithEvents lbDocNO As System.Windows.Forms.Label
    Friend WithEvents lbStock As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lbDetail As System.Windows.Forms.Label
    Friend WithEvents lbConID As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents numQty2 As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkDocState As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents lbDetl2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboLoop As System.Windows.Forms.ComboBox
    Friend WithEvents chkLoop As System.Windows.Forms.CheckBox
    Friend WithEvents cboCLoop As System.Windows.Forms.ComboBox
    Friend WithEvents chkCanN As System.Windows.Forms.CheckBox
    Friend WithEvents chkOrderLoop As System.Windows.Forms.CheckBox
    Friend WithEvents chkOrderCar As System.Windows.Forms.CheckBox
    Friend WithEvents lbState1 As System.Windows.Forms.Label
    Friend WithEvents lbState As System.Windows.Forms.Label
    Friend WithEvents lbServerDB As System.Windows.Forms.Label


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
    Friend WithEvents lsvShowBill As System.Windows.Forms.ListView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkDate As System.Windows.Forms.CheckBox
    Friend WithEvents Date01 As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbOK As System.Windows.Forms.Button
    Friend WithEvents chkTransport As System.Windows.Forms.CheckBox
    Friend WithEvents cboTransport As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lbSumFac As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkOrderLoop = New System.Windows.Forms.CheckBox()
        Me.chkOrderCar = New System.Windows.Forms.CheckBox()
        Me.lsvShowBill = New System.Windows.Forms.ListView()
        Me.lbSumFac = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lbState = New System.Windows.Forms.Label()
        Me.lbState1 = New System.Windows.Forms.Label()
        Me.lbDetl2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lbDetail = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.numQty2 = New System.Windows.Forms.NumericUpDown()
        Me.lbConID = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.lbStock = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lbDocNO = New System.Windows.Forms.Label()
        Me.lbArCode = New System.Windows.Forms.Label()
        Me.lbStkCode = New System.Windows.Forms.Label()
        Me.chkCheck = New System.Windows.Forms.CheckBox()
        Me.txtQty2 = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lbQty1 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lbStkName = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Date01 = New System.Windows.Forms.DateTimePicker()
        Me.cboTransport = New System.Windows.Forms.ComboBox()
        Me.chkTransport = New System.Windows.Forms.CheckBox()
        Me.chkDate = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmbOK = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.cboLoop = New System.Windows.Forms.ComboBox()
        Me.chkLoop = New System.Windows.Forms.CheckBox()
        Me.cboCLoop = New System.Windows.Forms.ComboBox()
        Me.chkCanN = New System.Windows.Forms.CheckBox()
        Me.chkDocState = New System.Windows.Forms.CheckBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.lbServerDB = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.numQty2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.DimGray
        Me.GroupBox1.Controls.Add(Me.chkOrderLoop)
        Me.GroupBox1.Controls.Add(Me.chkOrderCar)
        Me.GroupBox1.Controls.Add(Me.lsvShowBill)
        Me.GroupBox1.Controls.Add(Me.lbSumFac)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Location = New System.Drawing.Point(2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1274, 414)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'chkOrderLoop
        '
        Me.chkOrderLoop.AutoSize = True
        Me.chkOrderLoop.Checked = True
        Me.chkOrderLoop.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOrderLoop.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkOrderLoop.ForeColor = System.Drawing.Color.Aqua
        Me.chkOrderLoop.Location = New System.Drawing.Point(10, 370)
        Me.chkOrderLoop.Name = "chkOrderLoop"
        Me.chkOrderLoop.Size = New System.Drawing.Size(153, 37)
        Me.chkOrderLoop.TabIndex = 41
        Me.chkOrderLoop.Text = "ลำดับ-รอบ"
        Me.chkOrderLoop.UseVisualStyleBackColor = True
        '
        'chkOrderCar
        '
        Me.chkOrderCar.AutoSize = True
        Me.chkOrderCar.Checked = True
        Me.chkOrderCar.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOrderCar.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkOrderCar.ForeColor = System.Drawing.Color.Chartreuse
        Me.chkOrderCar.Location = New System.Drawing.Point(183, 370)
        Me.chkOrderCar.Name = "chkOrderCar"
        Me.chkOrderCar.Size = New System.Drawing.Size(136, 37)
        Me.chkOrderCar.TabIndex = 40
        Me.chkOrderCar.Text = "ลำดับ-รถ"
        Me.chkOrderCar.UseVisualStyleBackColor = True
        '
        'lsvShowBill
        '
        Me.lsvShowBill.BackColor = System.Drawing.Color.DimGray
        Me.lsvShowBill.Font = New System.Drawing.Font("Tahoma", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lsvShowBill.ForeColor = System.Drawing.Color.Snow
        Me.lsvShowBill.FullRowSelect = True
        Me.lsvShowBill.GridLines = True
        Me.lsvShowBill.HideSelection = False
        Me.lsvShowBill.Location = New System.Drawing.Point(6, 13)
        Me.lsvShowBill.Name = "lsvShowBill"
        Me.lsvShowBill.Size = New System.Drawing.Size(1256, 351)
        Me.lsvShowBill.TabIndex = 12
        Me.lsvShowBill.UseCompatibleStateImageBehavior = False
        '
        'lbSumFac
        '
        Me.lbSumFac.BackColor = System.Drawing.Color.PaleGoldenrod
        Me.lbSumFac.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbSumFac.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lbSumFac.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbSumFac.Location = New System.Drawing.Point(1105, 372)
        Me.lbSumFac.Name = "lbSumFac"
        Me.lbSumFac.Size = New System.Drawing.Size(119, 30)
        Me.lbSumFac.TabIndex = 37
        Me.lbSumFac.Text = "0.0"
        Me.lbSumFac.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label3.Location = New System.Drawing.Point(1222, 382)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 20)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "Ton."
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("MS Reference Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Yellow
        Me.Label4.Location = New System.Drawing.Point(1037, 376)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 24)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "นน. รวม :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.DimGray
        Me.GroupBox2.Controls.Add(Me.lbState)
        Me.GroupBox2.Controls.Add(Me.lbState1)
        Me.GroupBox2.Controls.Add(Me.lbDetl2)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.lbDetail)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.numQty2)
        Me.GroupBox2.Controls.Add(Me.lbConID)
        Me.GroupBox2.Controls.Add(Me.cmdOK)
        Me.GroupBox2.Controls.Add(Me.lbStock)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.lbDocNO)
        Me.GroupBox2.Controls.Add(Me.lbArCode)
        Me.GroupBox2.Controls.Add(Me.lbStkCode)
        Me.GroupBox2.Controls.Add(Me.chkCheck)
        Me.GroupBox2.Controls.Add(Me.txtQty2)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.lbQty1)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.lbStkName)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(4, 7)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1245, 121)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        '
        'lbState
        '
        Me.lbState.AutoSize = True
        Me.lbState.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbState.ForeColor = System.Drawing.Color.Yellow
        Me.lbState.Location = New System.Drawing.Point(716, 86)
        Me.lbState.Name = "lbState"
        Me.lbState.Size = New System.Drawing.Size(19, 19)
        Me.lbState.TabIndex = 54
        Me.lbState.Text = "0"
        '
        'lbState1
        '
        Me.lbState1.AutoSize = True
        Me.lbState1.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbState1.ForeColor = System.Drawing.Color.Yellow
        Me.lbState1.Location = New System.Drawing.Point(765, 86)
        Me.lbState1.Name = "lbState1"
        Me.lbState1.Size = New System.Drawing.Size(111, 19)
        Me.lbState1.TabIndex = 53
        Me.lbState1.Text = "สถานะ เปิดบิล"
        '
        'lbDetl2
        '
        Me.lbDetl2.BackColor = System.Drawing.Color.White
        Me.lbDetl2.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbDetl2.ForeColor = System.Drawing.Color.Maroon
        Me.lbDetl2.Location = New System.Drawing.Point(44, 83)
        Me.lbDetl2.Name = "lbDetl2"
        Me.lbDetl2.Size = New System.Drawing.Size(533, 30)
        Me.lbDetl2.TabIndex = 50
        Me.lbDetl2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label7.Location = New System.Drawing.Point(19, 91)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(51, 23)
        Me.Label7.TabIndex = 52
        Me.Label7.Text = "#"
        '
        'lbDetail
        '
        Me.lbDetail.BackColor = System.Drawing.Color.White
        Me.lbDetail.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbDetail.ForeColor = System.Drawing.Color.Maroon
        Me.lbDetail.Location = New System.Drawing.Point(43, 49)
        Me.lbDetail.Name = "lbDetail"
        Me.lbDetail.Size = New System.Drawing.Size(534, 30)
        Me.lbDetail.TabIndex = 49
        Me.lbDetail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label2.Location = New System.Drawing.Point(9, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 23)
        Me.Label2.TabIndex = 51
        Me.Label2.Text = "con."
        '
        'numQty2
        '
        Me.numQty2.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.numQty2.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.numQty2.ForeColor = System.Drawing.Color.Yellow
        Me.numQty2.Location = New System.Drawing.Point(893, 25)
        Me.numQty2.Maximum = New Decimal(New Integer() {9999, 0, 0, 0})
        Me.numQty2.Name = "numQty2"
        Me.numQty2.Size = New System.Drawing.Size(102, 40)
        Me.numQty2.TabIndex = 14
        '
        'lbConID
        '
        Me.lbConID.BackColor = System.Drawing.Color.Gray
        Me.lbConID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbConID.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbConID.ForeColor = System.Drawing.Color.Yellow
        Me.lbConID.Location = New System.Drawing.Point(443, 48)
        Me.lbConID.Name = "lbConID"
        Me.lbConID.Size = New System.Drawing.Size(81, 27)
        Me.lbConID.TabIndex = 49
        Me.lbConID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.DarkKhaki
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOK.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Black
        Me.cmdOK.Location = New System.Drawing.Point(1105, 20)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(127, 60)
        Me.cmdOK.TabIndex = 43
        Me.cmdOK.Text = "บันทึก"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'lbStock
        '
        Me.lbStock.BackColor = System.Drawing.Color.Black
        Me.lbStock.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbStock.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbStock.ForeColor = System.Drawing.Color.White
        Me.lbStock.Location = New System.Drawing.Point(587, 22)
        Me.lbStock.Name = "lbStock"
        Me.lbStock.Size = New System.Drawing.Size(106, 44)
        Me.lbStock.TabIndex = 47
        Me.lbStock.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label5.Location = New System.Drawing.Point(530, 52)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 23)
        Me.Label5.TabIndex = 48
        Me.Label5.Text = "Stock"
        '
        'lbDocNO
        '
        Me.lbDocNO.BackColor = System.Drawing.Color.Gray
        Me.lbDocNO.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbDocNO.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbDocNO.ForeColor = System.Drawing.Color.Yellow
        Me.lbDocNO.Location = New System.Drawing.Point(295, 48)
        Me.lbDocNO.Name = "lbDocNO"
        Me.lbDocNO.Size = New System.Drawing.Size(148, 27)
        Me.lbDocNO.TabIndex = 46
        Me.lbDocNO.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lbArCode
        '
        Me.lbArCode.BackColor = System.Drawing.Color.Gray
        Me.lbArCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbArCode.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbArCode.ForeColor = System.Drawing.Color.Yellow
        Me.lbArCode.Location = New System.Drawing.Point(240, 48)
        Me.lbArCode.Name = "lbArCode"
        Me.lbArCode.Size = New System.Drawing.Size(62, 27)
        Me.lbArCode.TabIndex = 45
        Me.lbArCode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lbStkCode
        '
        Me.lbStkCode.BackColor = System.Drawing.Color.Gray
        Me.lbStkCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbStkCode.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbStkCode.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbStkCode.Location = New System.Drawing.Point(42, 48)
        Me.lbStkCode.Name = "lbStkCode"
        Me.lbStkCode.Size = New System.Drawing.Size(201, 27)
        Me.lbStkCode.TabIndex = 44
        Me.lbStkCode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkCheck
        '
        Me.chkCheck.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.chkCheck.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkCheck.ForeColor = System.Drawing.Color.Beige
        Me.chkCheck.Location = New System.Drawing.Point(1005, 38)
        Me.chkCheck.Name = "chkCheck"
        Me.chkCheck.Size = New System.Drawing.Size(95, 27)
        Me.chkCheck.TabIndex = 37
        Me.chkCheck.Text = "เรียบร้อย"
        Me.chkCheck.UseVisualStyleBackColor = False
        '
        'txtQty2
        '
        Me.txtQty2.BackColor = System.Drawing.SystemColors.HotTrack
        Me.txtQty2.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtQty2.ForeColor = System.Drawing.Color.White
        Me.txtQty2.Location = New System.Drawing.Point(894, 26)
        Me.txtQty2.Name = "txtQty2"
        Me.txtQty2.Size = New System.Drawing.Size(102, 40)
        Me.txtQty2.TabIndex = 6
        Me.txtQty2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label8.Location = New System.Drawing.Point(837, 46)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(51, 23)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "จัดได้"
        '
        'lbQty1
        '
        Me.lbQty1.BackColor = System.Drawing.Color.Yellow
        Me.lbQty1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbQty1.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbQty1.Location = New System.Drawing.Point(732, 25)
        Me.lbQty1.Name = "lbQty1"
        Me.lbQty1.Size = New System.Drawing.Size(99, 40)
        Me.lbQty1.TabIndex = 2
        Me.lbQty1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label6.Location = New System.Drawing.Point(695, 46)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 23)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "เบิก"
        '
        'lbStkName
        '
        Me.lbStkName.BackColor = System.Drawing.Color.YellowGreen
        Me.lbStkName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbStkName.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbStkName.Location = New System.Drawing.Point(42, 19)
        Me.lbStkName.Name = "lbStkName"
        Me.lbStkName.Size = New System.Drawing.Size(535, 27)
        Me.lbStkName.TabIndex = 1
        Me.lbStkName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label1.Location = New System.Drawing.Point(13, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ขื่อ"
        '
        'Date01
        '
        Me.Date01.CalendarFont = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Date01.CalendarForeColor = System.Drawing.Color.Red
        Me.Date01.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Date01.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.Date01.Location = New System.Drawing.Point(59, 23)
        Me.Date01.Name = "Date01"
        Me.Date01.Size = New System.Drawing.Size(158, 36)
        Me.Date01.TabIndex = 7
        '
        'cboTransport
        '
        Me.cboTransport.BackColor = System.Drawing.Color.SteelBlue
        Me.cboTransport.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cboTransport.ForeColor = System.Drawing.SystemColors.Info
        Me.cboTransport.Location = New System.Drawing.Point(286, 20)
        Me.cboTransport.Name = "cboTransport"
        Me.cboTransport.Size = New System.Drawing.Size(151, 41)
        Me.cboTransport.TabIndex = 36
        '
        'chkTransport
        '
        Me.chkTransport.BackColor = System.Drawing.Color.DimGray
        Me.chkTransport.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkTransport.ForeColor = System.Drawing.Color.Beige
        Me.chkTransport.Location = New System.Drawing.Point(223, 36)
        Me.chkTransport.Name = "chkTransport"
        Me.chkTransport.Size = New System.Drawing.Size(88, 27)
        Me.chkTransport.TabIndex = 12
        Me.chkTransport.Text = "ลำดับเอกสาร"
        Me.chkTransport.UseVisualStyleBackColor = False
        '
        'chkDate
        '
        Me.chkDate.BackColor = System.Drawing.Color.DimGray
        Me.chkDate.Checked = True
        Me.chkDate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDate.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkDate.ForeColor = System.Drawing.Color.Beige
        Me.chkDate.Location = New System.Drawing.Point(9, 38)
        Me.chkDate.Name = "chkDate"
        Me.chkDate.Size = New System.Drawing.Size(55, 24)
        Me.chkDate.TabIndex = 10
        Me.chkDate.Text = "วันที่"
        Me.chkDate.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Maroon
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(1110, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(124, 53)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "จบการทำงาน"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'cmbOK
        '
        Me.cmbOK.BackColor = System.Drawing.Color.YellowGreen
        Me.cmbOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbOK.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cmbOK.ForeColor = System.Drawing.Color.Black
        Me.cmbOK.Location = New System.Drawing.Point(980, 14)
        Me.cmbOK.Name = "cmbOK"
        Me.cmbOK.Size = New System.Drawing.Size(124, 53)
        Me.cmbOK.TabIndex = 11
        Me.cmbOK.Text = "RUN"
        Me.cmbOK.UseVisualStyleBackColor = False
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.DimGray
        Me.GroupBox3.Controls.Add(Me.cboLoop)
        Me.GroupBox3.Controls.Add(Me.chkLoop)
        Me.GroupBox3.Controls.Add(Me.cboCLoop)
        Me.GroupBox3.Controls.Add(Me.cmbOK)
        Me.GroupBox3.Controls.Add(Me.chkCanN)
        Me.GroupBox3.Controls.Add(Me.chkDocState)
        Me.GroupBox3.Controls.Add(Me.Date01)
        Me.GroupBox3.Controls.Add(Me.Button1)
        Me.GroupBox3.Controls.Add(Me.chkDate)
        Me.GroupBox3.Controls.Add(Me.cboTransport)
        Me.GroupBox3.Controls.Add(Me.chkTransport)
        Me.GroupBox3.Location = New System.Drawing.Point(4, 131)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1245, 77)
        Me.GroupBox3.TabIndex = 14
        Me.GroupBox3.TabStop = False
        '
        'cboLoop
        '
        Me.cboLoop.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cboLoop.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cboLoop.ForeColor = System.Drawing.SystemColors.Info
        Me.cboLoop.Location = New System.Drawing.Point(656, 20)
        Me.cboLoop.Name = "cboLoop"
        Me.cboLoop.Size = New System.Drawing.Size(127, 41)
        Me.cboLoop.TabIndex = 85
        '
        'chkLoop
        '
        Me.chkLoop.BackColor = System.Drawing.Color.Transparent
        Me.chkLoop.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkLoop.ForeColor = System.Drawing.Color.GreenYellow
        Me.chkLoop.Location = New System.Drawing.Point(588, 35)
        Me.chkLoop.Name = "chkLoop"
        Me.chkLoop.Size = New System.Drawing.Size(88, 27)
        Me.chkLoop.TabIndex = 87
        Me.chkLoop.Text = "รอบที่"
        Me.chkLoop.UseVisualStyleBackColor = False
        '
        'cboCLoop
        '
        Me.cboCLoop.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.cboCLoop.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cboCLoop.ForeColor = System.Drawing.SystemColors.Info
        Me.cboCLoop.Location = New System.Drawing.Point(850, 20)
        Me.cboCLoop.Name = "cboCLoop"
        Me.cboCLoop.Size = New System.Drawing.Size(124, 41)
        Me.cboCLoop.TabIndex = 84
        '
        'chkCanN
        '
        Me.chkCanN.BackColor = System.Drawing.Color.Transparent
        Me.chkCanN.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkCanN.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.chkCanN.Location = New System.Drawing.Point(791, 34)
        Me.chkCanN.Name = "chkCanN"
        Me.chkCanN.Size = New System.Drawing.Size(75, 27)
        Me.chkCanN.TabIndex = 86
        Me.chkCanN.Text = "คันที"
        Me.chkCanN.UseVisualStyleBackColor = False
        '
        'chkDocState
        '
        Me.chkDocState.BackColor = System.Drawing.Color.DimGray
        Me.chkDocState.Checked = True
        Me.chkDocState.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDocState.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkDocState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkDocState.Location = New System.Drawing.Point(476, 35)
        Me.chkDocState.Name = "chkDocState"
        Me.chkDocState.Size = New System.Drawing.Size(106, 24)
        Me.chkDocState.TabIndex = 43
        Me.chkDocState.Text = "ยังไม่เปิดบิล"
        Me.chkDocState.UseVisualStyleBackColor = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(2, 422)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1269, 245)
        Me.TabControl1.TabIndex = 14
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox3)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1261, 219)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "TabPage1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1261, 219)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "TabPage2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'lbServerDB
        '
        Me.lbServerDB.BackColor = System.Drawing.Color.DarkRed
        Me.lbServerDB.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.lbServerDB.Location = New System.Drawing.Point(1, 669)
        Me.lbServerDB.Name = "lbServerDB"
        Me.lbServerDB.Size = New System.Drawing.Size(1260, 24)
        Me.lbServerDB.TabIndex = 15
        Me.lbServerDB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmBegin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(1292, 711)
        Me.Controls.Add(Me.lbServerDB)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Name = "frmBegin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "เบิกสินค้า  (updte 01-09-58)"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.numQty2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region









    Sub HeadList()

        If chkTB = True Then

            lsvShowBill.Columns.Clear()
            lsvShowBill.Items.Clear()
            chkTB = False

        End If

        lsvShowBill.Columns.Add("เลขที่", 110, HorizontalAlignment.Left) '1
        lsvShowBill.Columns.Add("#", 1, HorizontalAlignment.Left)  '2
        lsvShowBill.Columns.Add("รอบ", 50, HorizontalAlignment.Center) '3
        lsvShowBill.Columns.Add("คัน", 50, HorizontalAlignment.Center) '4

        'lsvShowBill.Columns.Add("ชื่อรถจัดส่ง", 0, HorizontalAlignment.Left) '3
        'lsvShowBill.Columns.Add("โซน", 0, HorizontalAlignment.Left) '4

        lsvShowBill.Columns.Add("ชื่อลูกค้า", 200, HorizontalAlignment.Left) '5
        'lsvShowBill.Columns.Add("Vatลูกค้า", 40, HorizontalAlignment.Right) '6
        lsvShowBill.Columns.Add("ชื่อสินค้า", 450, HorizontalAlignment.Left) '7
        lsvShowBill.Columns.Add("Stock", 80, HorizontalAlignment.Right) '8
        lsvShowBill.Columns.Add("เบิก", 80, HorizontalAlignment.Right) '9
        lsvShowBill.Columns.Add("จัด", 80, HorizontalAlignment.Right) '10

        'lsvShowBill.Columns.Add("ที่อยู่", 0, HorizontalAlignment.Left) '
        lsvShowBill.Columns.Add("Condition", 350, HorizontalAlignment.Left) '11

        lsvShowBill.Columns.Add("SalesMan", 100, HorizontalAlignment.Left) '12
        lsvShowBill.Columns.Add("รหัสสินค้า", 100, HorizontalAlignment.Left) '13
        lsvShowBill.Columns.Add("รหัสลูกค้า", 100, HorizontalAlignment.Left) '14
        lsvShowBill.Columns.Add("state", 50, HorizontalAlignment.Left) '15  จัดแล้ว
        lsvShowBill.Columns.Add("เลขที่ Order", 50, HorizontalAlignment.Center) '16
        lsvShowBill.Columns.Add("หมายเหตุ", 150, HorizontalAlignment.Left) '17
        lsvShowBill.Columns.Add("เปิดบิล", 150, HorizontalAlignment.Left) '18
        lsvShowBill.View = View.Details
        lsvShowBill.GridLines = True


        chkTB = True

    End Sub

    Private Sub frmBegin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call DBtools.openDB()
        'Me.Text = Me.Text & " หมายเลข #DataBase Server =" & DBConnString.strConn2
        lbServerDB.Text = " หมายเลข #DataBase Server =" & DBConnString.strConn2
        Call comboTran()
        Call listCarLoop()
        Call listLoop()
        chkLoad = True
        'chkItem = False
    End Sub

    Sub comboTran()
        Dim DS As DataSet = New DataSet
        Dim DA As SqlClient.SqlDataAdapter
        DS.Clear()
        txtSQL = "Select  * "
        txtSQL = txtSQL & " From transportation "
        DA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        DA.Fill(DS, "Car")
        If DS.Tables("Car").Rows.Count > 0 Then
            With cboTransport
                .DataSource = DS.Tables("Car")
                .DisplayMember = "TranName"
                .ValueMember = "TranCode"
            End With
        End If
    End Sub

    Private Sub cmbOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOK.Click
        Call HeadList()
        Call Search()
        Call txtCLS()
    End Sub

    Sub Search()

        'Dim CusName2 As String
        'Dim CusName3 As String
        Dim SumFactor As Double = 0.0
        Dim SumFactor1 As Double = 0.0

        Dim sumAmt As Double = 0
        Dim chkZoneA As String = ""
        Dim chkZoneB As String = ""
        Dim strTran As String = ""
        Dim strCarNum As String = ""
        Dim chkState0 As Integer = 0
        Dim chkColor As Integer = 0
        SumFactor = 0.0
        SumFactor1 = 0.0

        lbSumFac.Text = 0
        'lbSumAmt.Text = 0

        If chkBill = True Then
            searchDS.Tables("BillData").Clear()
            chkBill = False
        End If

        '======================   เลือกข้อมูลใบเบิก  ============================

        txtSQL1 = " SELECt TranDataH.Trh_Transport,CarMast.Car_Contact_Name,"
        txtSQL1 = txtSQL1 & "TranDataH.Trh_Bill,TranDataH.Trh_RunID,"

        txtSQL1 = txtSQL1 & "CarMast.Car_Name, TranDataD.Dtl_type, TranDataD.Dtl_date,TranDataD.Dtl_dueDate,"
        txtSQL1 = txtSQL1 & " TranDataD.Dtl_no,TranDataD.Dtl_Chk,TranDataD.Dtl_State, TranDataD.Dtl_con_id, ArFile.AR_Cus_iD,ArFile.AR_NAME,"
        'txtSQL1 = txtSQL1 & "(ArFile.Ar_Addr + ' ' + ArFile.Ar_Addr_1+ ' ' + ArFile.Ar_Addr_2)as Ar_Addr,"

        txtSQL1 = txtSQL1 & "TranDataD.Dtl_detail,TRanDataD.Dtl_Condition,SalesMan.SL_Name, "
        txtSQL1 = txtSQL1 & "(Ar_Map_Code + '-' + Map_S_Line_name)as Zone,"
        txtSQL1 = txtSQL1 & " left(Ar_Map_Code,1)as Map_Grp,"
        txtSQL1 = txtSQL1 & " TranDataD.Dtl_idtrade,BaseMast.Stk_Name_1,BaseMast.Stk_Factor,StkDetl.Dtl_Bal_Q1,"
        txtSQL1 = txtSQL1 & " TranDataD.Dtl_price, TranDataD.Dtl_num,  TranDataD.Dtl_sum,  "
        txtSQL1 = txtSQL1 & " (BaseMast.Stk_Factor * TranDataD.Dtl_num) AS FacSum, "
        txtSQL1 = txtSQL1 & " (TranDataD.Dtl_price / BaseMast.Stk_Factor ) AS PrPerKg,  "
        txtSQL1 = txtSQL1 & " TranDataD.dtl_num_2,"
        txtSQL1 = txtSQL1 & " (TranDataD.Dtl_price * TranDataD.Dtl_Num ) AS PrPerB "


        txtSQL1 = txtSQL1 & "FROM  TranDataD RIGHT JOIN TranDataH  "
        txtSQL1 = txtSQL1 & "ON TranDataD.Dtl_no = TranDataH.Trh_No "
        txtSQL1 = txtSQL1 & "AND TranDataD.Dtl_type = TranDataH.Trh_Type "

        txtSQL1 = txtSQL1 & "LEFT JOIN ArFile "
        txtSQL1 = txtSQL1 & "ON TranDataD.Dtl_idcus = ArFile.AR_CUS_ID  "
        txtSQL1 = txtSQL1 & " Left Join MapSLineMast "
        txtSQL1 = txtSQL1 & " On Ar_Map_Code=Map_S_Line_Code "

        txtSQL1 = txtSQL1 & "LEFT JOIN BaseMast "
        txtSQL1 = txtSQL1 & "ON TranDataD.Dtl_idtrade = BaseMast.Stk_Code "
        txtSQL1 = txtSQL1 & "Left Join StkDetl "
        txtSQL1 = txtSQL1 & "On BaseMast.Stk_Code=StkDetl.Dtl_Code "
        txtSQL1 = txtSQL1 & "Left Join CarMast "
        txtSQL1 = txtSQL1 & "On TranDataH.Trh_Cus=CarMast.Car_No "

        txtSQL1 = txtSQL1 & "Left Join SalesMAN "
        txtSQL1 = txtSQL1 & "On Arfile.Ar_Sales=SalesMan.SL_ID "


        txtSQL1 = txtSQL1 & "WHERE (TranDataD.Dtl_type = 'P') "
        txtSQL1 = txtSQL1 & "And (Stkdetl.Dtl_WH='01') "
        If chkDocState.Checked = True Then
            '  ถ้าเป็น True  เลือกเฉพาะที่ยังไม่เปิดบิล ถ้าเป็น False เลือกทั้งหมด
            txtSQL1 = txtSQL1 & "And (TranDataD.Dtl_State='0' ) " ' เช็ค State ว่าได้เปิดบิลไปหรือยัง ถ้าเปิดบิลแล้วจะเป็น 1
        Else
            'txtSQL1 = txtSQL1 & "And (TranDataD.Dtl_State='0' ) " ' เช็ค State ว่าได้เปิดบิลไปหรือยัง ถ้าเปิดบิลแล้วจะเป็น 1
        End If


        If chkDate.Checked = True Then
            txtSQL1 = txtSQL1 & "And (TranDataD.Dtl_dueDate= '" & (Format(DateAdd(DateInterval.Year, -543, Date01.Value), "MM/dd/yyyy")) & "')  "
            'txtSQL1 = txtSQL1 & "And (TranDataD.Dtl_date<= '" & (Format(DateAdd(DateInterval.Year, -543, Date02.Value), "MM/dd/yyyy")) & "') "
        End If

        If chkLoop.Checked = True Then
            txtSQL1 = txtSQL1 & " and TranDatah.trh_bill= '" & cboLoop.SelectedValue & "'"
        End If
        If chkCanN.Checked = True Then
            txtSQL1 = txtSQL1 & " and TranDatah.trh_RunID= '" & cboCLoop.SelectedValue & "'"
        End If

        If chkTransport.Checked = True Then
            txtSQL1 = txtSQL1 & " and TranDatah.trh_transport= '" & cboTransport.SelectedValue & "'"
        End If
        If chkOrderCar.Checked = True Then
            txtSQL1 = txtSQL1 & "Order by tranDataH.Trh_Bill desc,TranDataH.Trh_RunID desc,dtl_no desc,TranDataH.Trh_Transport,dtl_idcus "
        ElseIf chkOrderLoop.Checked = True Then
            txtSQL1 = txtSQL1 & "Order by tranDataH.Trh_Bill desc,TranDataH.Trh_RunID desc,dtl_no desc,TranDataH.Trh_Transport,dtl_idcus "
        Else
            txtSQL1 = txtSQL1 & " Order by dtl_no desc,TranDataH.Trh_Transport,dtl_idcus "
        End If


        searchDA = New SqlClient.SqlDataAdapter(txtSQL1, Conn)
        searchDA.Fill(searchDS, "BillData")
        chkBill = True

        Dim maxI As Integer
        Dim i As Integer


        maxI = searchDS.Tables("BillData").Rows.Count
        chkZoneA = ""
        chkRow = True

        For i = 0 To maxI - 1 'วนลูปตามเลขที่order

            '=========================================================
            '=========================================================

            BillDate = searchDS.Tables("BillData").Rows(i).Item("Dtl_dueDate")
            'docSelDate = searchDS.Tables("BillData").Rows(i).Item("Dtl_Date")
            BNo = searchDS.Tables("BillData").Rows(i).Item("Dtl_no")
            'strDocCon = searchDS.Tables("BillData").Rows(i).Item("Dtl_Con_ID")

            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("Dtl_chk")) Then
                intChk = 0
                'chkColor = 2
            Else
                If searchDS.Tables("BillData").Rows(i).Item("Dtl_Chk") = 0 Then
                    intChk = 0
                    ' chkColor = 2
               
                ElseIf searchDS.Tables("BillData").Rows(i).Item("Dtl_Chk") = 1 Then
                    intChk = 1
                    'chkColor = 0
                End If

            End If



            If searchDS.Tables("BillData").Rows.Count - 1 = i Then
                BNo2 = searchDS.Tables("BillData").Rows(i).Item("Dtl_no")
            Else
                BNo2 = searchDS.Tables("BillData").Rows(i + 1).Item("Dtl_no")
            End If

            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("Dtl_State")) Then
                chkState0 = 0
            Else
                chkState0 = searchDS.Tables("BillData").Rows(i).Item("Dtl_state")
            End If


            ONo = searchDS.Tables("BillData").Rows(i).Item("Dtl_con_id")
            StkID = searchDS.Tables("BillData").Rows(i).Item("Dtl_idtrade")
            stkStock = searchDS.Tables("BillData").Rows(i).Item("Dtl_Bal_Q1")

            StkName = DBtools.getStkName(StkID) 'searchDS.Tables("BillData").Rows(i).Item("Stk_Name_1")
            'StkPrice = searchDS.Tables("BillData").Rows(i).Item("Dtl_price")
            StkPrPerKg = searchDS.Tables("BillData").Rows(i).Item("PrPerKg")

            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("Zone")) Then
                Zone = "ยังไม่ได้ระบุ"
            Else
                Zone = searchDS.Tables("BillData").Rows(i).Item("Zone")
            End If
            '==============  ตรวจสอบการเปิดบิล
            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("Dtl_State")) Then
                lbState.Text = 0
            ElseIf (searchDS.Tables("BillData").Rows(i).Item("Dtl_State") = 0) Then
                lbState.Text = 0
            ElseIf (searchDS.Tables("BillData").Rows(i).Item("Dtl_State") = 1) Then 'เปิดบิลแล้ว
                lbState.Text = 1
            End If

            CusName = searchDS.Tables("BillData").Rows(i).Item("ar_name")
            cusCode = searchDS.Tables("BillData").Rows(i).Item("ar_Cus_ID")
            BNum = searchDS.Tables("BillData").Rows(i).Item("dtl_num")


            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("dtl_num_2")) Then
                NumQ2 = 0
            Else
                NumQ2 = searchDS.Tables("BillData").Rows(i).Item("dtl_num_2")
            End If




            BPrice = searchDS.Tables("BillData").Rows(i).Item("dtl_price")
            BSum = searchDS.Tables("BillData").Rows(i).Item("dtl_sum")
            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("Car_Name")) Then
                Transport = "ไม่พบข้อมูล"
            Else
                Transport = searchDS.Tables("BillData").Rows(i).Item("Car_Name")
            End If


            'BBill = searchDS.Tables("BillData").Rows(i).Item("trh_bill")
            SumFactor = searchDS.Tables("BillData").Rows(i).Item("FacSum")

            'If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("Ar_Addr")) Then
            '    cusAddr = "ไม่พบข้อมูล"
            'Else
            '    cusAddr = searchDS.Tables("BillData").Rows(i).Item("Ar_Addr")
            'End If
            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("Dtl_condition")) Then
                cusDetail = "ไม่พบข้อมูล"
            Else
                cusDetail = searchDS.Tables("BillData").Rows(i).Item("Dtl_condition")
            End If
            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("Dtl_Detail")) Then
                cusDetail2 = "ไม่พบข้อมูล"
            Else
                cusDetail2 = searchDS.Tables("BillData").Rows(i).Item("Dtl_Detail")
            End If

            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("trh_Bill")) Then
                strTran = 0
            Else
                strTran = searchDS.Tables("BillData").Rows(i).Item("trh_Bill")
            End If

            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("trh_RunID")) Then
                strCarNum = 0
            Else
                strCarNum = searchDS.Tables("BillData").Rows(i).Item("trh_RunID")
            End If

            If IsDBNull(searchDS.Tables("BillData").Rows(i).Item("SL_Name")) Then
                cusSales = "ไม่พบข้อมูล"
            Else
                cusSales = searchDS.Tables("BillData").Rows(i).Item("SL_Name")
            End If


            'AnyData = New String() {BNo, BillDate.ToString("dd/MM/yy"), Trim(Transport), Zone, CusName, StkName, BNum.ToString("#,##0.00"), SumFactor.ToString("#,##0.00"), cusAddr, cusDetail, cusSales}
            AnyData = New String() {BNo, BillDate.ToString("dd/MM/yy"), strTran, strCarNum, CusName, StkName, stkStock.ToString("#,##0"), BNum.ToString("#,##0"), NumQ2.ToString("#,##0"), cusDetail, cusSales, StkID, cusCode, intChk, ONo, cusDetail2, chkState0}
            sumW = BSum + sumW


            lvi = New ListViewItem(AnyData)
            lsvShowBill.Items.Add(lvi)
            chkZoneB = BNo

            SumFactor1 = SumFactor1 + SumFactor
            sumAmt = sumAmt + searchDS.Tables("BillData").Rows(i).Item("PrPerB")

            Call colorList(intChk, chkState0)

            'If intChk = 0 Then

            'ElseIf intChk = 1 Then
            '    Call colorList(1)
            'ElseIf intChk = 2 Then
            '    Call colorList(2)
            'End If

        Next
        'lbSumAmt.Text = Format(sumAmt, "#,##0.00")
        lbSumFac.Text = Format(SumFactor1 / 1000, "#,##0.000")
        'lbPrAVG.Text = Format(sumAmt / SumFactor1, "#,##0.000")

    End Sub
    Sub listLoop()

        Dim dTBL As New DataTable
        Dim dRowL As DataRow
        'Call DBTools.connDB()

        dTBL.Columns.Add(New DataColumn("L_Type", GetType(String)))
        dTBL.Columns.Add(New DataColumn("L_Name", GetType(String)))

        dRowL = dTBL.NewRow
        dRowL(0) = "1"
        dRowL(1) = "รอบที่ 1"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "2"
        dRowL(1) = "รอบที่ 2"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "3"
        dRowL(1) = "รอบที่ 3"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "4"
        dRowL(1) = "รอบที่ 4"
        dTBL.Rows.Add(dRowL)

        With cboLoop
            .DataSource = dTBL
            .DisplayMember = "L_Name"
            .ValueMember = "L_Type"
        End With
        cboLoop.SelectedValue = 1

    End Sub

    Sub listCarLoop()

        Dim dTBL As New DataTable
        Dim dRowL As DataRow
        'Call DBTools.connDB()

        dTBL.Columns.Add(New DataColumn("C_Type", GetType(String)))
        dTBL.Columns.Add(New DataColumn("C_Name", GetType(String)))

        dRowL = dTBL.NewRow
        dRowL(0) = "1"
        dRowL(1) = "คันที่ 1"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "2"
        dRowL(1) = "คันที่ 2"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "3"
        dRowL(1) = "คันที่ 3"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "4"
        dRowL(1) = "คันที่ 4"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "5"
        dRowL(1) = "คันที่ 5"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "6"
        dRowL(1) = "คันที่ 6"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "7"
        dRowL(1) = "คันที่ 7"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "8"
        dRowL(1) = "คันที่ 8"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "9"
        dRowL(1) = "คันที่ 9"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "10"
        dRowL(1) = "คันที่ 10"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "11"
        dRowL(1) = "คันที่ 11"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "12"
        dRowL(1) = "คันที่ 12"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "13"
        dRowL(1) = "คันที่ 13"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "14"
        dRowL(1) = "คันที่ 14"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "15"
        dRowL(1) = "คันที่ 15"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "16"
        dRowL(1) = "คันที่ 16"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "17"
        dRowL(1) = "คันที่ 17"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "18"
        dRowL(1) = "คันที่ 18"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "19"
        dRowL(1) = "คันที่ 19"
        dTBL.Rows.Add(dRowL)

        dRowL = dTBL.NewRow
        dRowL(0) = "20"
        dRowL(1) = "คันที่ 20"
        dTBL.Rows.Add(dRowL)

        With cboCLoop
            .DataSource = dTBL
            .DisplayMember = "C_Name"
            .ValueMember = "C_Type"
        End With
        cboCLoop.SelectedValue = 1

    End Sub
    Sub colorList(ByVal chkCheck As Integer, chkState As Integer)

        If chkCheck = 1 And chkState = 1 Then
            With lvi
                .BackColor = System.Drawing.Color.YellowGreen
                .ForeColor = System.Drawing.Color.Black
                .UseItemStyleForSubItems = True

            End With
        ElseIf chkCheck = 0 And chkState = 1 Then
            With lvi
                .BackColor = System.Drawing.Color.DarkViolet
                .ForeColor = System.Drawing.Color.White
                .UseItemStyleForSubItems = True

            End With

        ElseIf chkState = 0 And chkCheck = 0 Then

            With lvi
                .BackColor = System.Drawing.Color.DarkOrange
                .ForeColor = System.Drawing.Color.Black
                .UseItemStyleForSubItems = True

            End With
        End If


    End Sub
    Private Sub frmBegin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            End
        End If
    End Sub

    Private Sub lsvShowBill_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lsvShowBill.Click

        Dim i As Integer
        ' Dim frm1 As New frm
        'Try
        chkCheck.Checked = False
        For i = 0 To lsvShowBill.SelectedItems.Count - 1

            lvi = lsvShowBill.SelectedItems(i)
            chkEditDoc = True
            chkNew = False
            EditNo = lsvShowBill.SelectedItems(i).SubItems.Item(0).Text()
            lbDocNO.Text = lsvShowBill.SelectedItems(i).SubItems.Item(0).Text()
            lbStock.Text = lsvShowBill.SelectedItems(i).SubItems.Item(6).Text()
            lbQty1.Text = lsvShowBill.SelectedItems(i).SubItems.Item(7).Text

            If lsvShowBill.SelectedItems(i).SubItems.Item(8).Text > 0 Then
                numQty2.Value = lsvShowBill.SelectedItems(i).SubItems.Item(8).Text
                txtQty2.Text = lsvShowBill.SelectedItems(i).SubItems.Item(8).Text
            Else
                'If IsDBNull(lsvShowBill.SelectedItems(i).SubItems.Item(5).Text) Then
                '    numQty2.Value = 0
                '    txtQty2.Text = 0
                'Else
                numQty2.Value = CDbl(lsvShowBill.SelectedItems(i).SubItems.Item(7).Text)
                txtQty2.Text = lsvShowBill.SelectedItems(i).SubItems.Item(7).Text
                'End If

            End If

            lbStkName.Text = getStkName(lsvShowBill.SelectedItems(i).SubItems.Item(11).Text)
            lbArCode.Text = lsvShowBill.SelectedItems(i).SubItems.Item(12).Text
            lbStkCode.Text = lsvShowBill.SelectedItems(i).SubItems.Item(11).Text
            lbDetail.Text = lsvShowBill.SelectedItems(i).SubItems.Item(9).Text
            lbDetl2.Text = lsvShowBill.SelectedItems(i).SubItems.Item(15).Text
            lbConID.Text = lsvShowBill.SelectedItems(i).SubItems.Item(14).Text
            chkCheck.Checked = lsvShowBill.SelectedItems(i).SubItems.Item(13).Text ' เช็คการตรวจนับขึ้นของ
            'if
            lbState.Text = lsvShowBill.SelectedItems(i).SubItems.Item(16).Text  ' เช็คการเปิดบิล
            If lbState.Text = 1 Then
                lbState1.Text = "สถานะ เปิดบิล"
                lbState1.BackColor = Color.GreenYellow
                lbState1.ForeColor = Color.Black
            Else
                lbState1.Text = "สถานะ ยังไม่เปิดบิล"
                lbState1.BackColor = Color.DarkViolet
                lbState1.ForeColor = Color.White
            End If

        Next

        Me.Hide()
        ' frm1.ShowDialog()
        Me.Show()
        'Call HeadList()
        'Call Search()
        chkNew = False
        chkEditDoc = False
        EditNo = ""
        'txtEditNo.Text = ""
        'lsvShowBill.Clear()
        'chkDate.Checked = False
        'Date01.Value = Now
        'Date02.Value = Now
        'Catch
        '    Exit Sub
        'End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        End

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        '  ให้เช็ค Dtl_Status
        'If getChkTypeP_State(lbDocNO.Text, lbStkCode.Text) = True Then
        '    '    'Call saveData_D()  ปลดออกในกรณีที่จะเอารายการที่จะเปิดบิลไปแล้วออก
        '    MsgBox("ใบเบิกนี้ได้มีการเปิดบิลแล้ว  ! ", MsgBoxStyle.Critical, "แจ้วเตือน")
        '    '    'Exit Sub

        'End If

        txtSQL = "Update TranDataD "
        txtSQL = txtSQL & "set dtl_num_2='" & numQty2.Value & "',"

        If chkCheck.Checked = True Then
            If lbState.Text = 0 Then
                txtSQL = txtSQL & "dtl_chk='1' "
            ElseIf lbState.Text = 1 Then
                txtSQL = txtSQL & "dtl_chk='2' "
            End If
        Else
            txtSQL = txtSQL & "dtl_chk='0' "
        End If

        txtSQL = txtSQL & "Where Dtl_type='P' "
        txtSQL = txtSQL & "And Dtl_Date='" & (Format(DateAdd(DateInterval.Year, -543, Date01.Value), "MM/dd/yyyy")) & "'"
        txtSQL = txtSQL & "And Dtl_no='" & lbDocNO.Text & "' "
        txtSQL = txtSQL & "And Dtl_idCus='" & lbArCode.Text & "' "
        txtSQL = txtSQL & "And Dtl_idTrade='" & lbStkCode.Text & "' "
        txtSQL = txtSQL & "And Dtl_Con_ID='" & lbConID.Text & "'"
        DBtools.dbSaveSQLsrv(txtSQL, "")

        Call HeadList()
        Call Search()
        Call txtCLS()
        'End If

    End Sub

    Sub txtCLS()

        txtQty2.Text = 0
        lbQty1.Text = 0
        lbStock.Text = 0
        numQty2.Value = 0
        chkCheck.Checked = False

        lbStkCode.Text = ""
        lbStkName.Text = ""
        lbArCode.Text = ""
        lbConID.Text = ""
        lbDocNO.Text = ""

    End Sub
    Private Sub txtQty2_DoubleClick(sender As Object, e As EventArgs) Handles txtQty2.DoubleClick
        Dim frmQTy2 As New frmQty2
        frmQTy2.ShowDialog()

    End Sub

    

    
  
    Private Sub chkDocState_CheckedChanged(sender As Object, e As EventArgs) Handles chkDocState.CheckedChanged
        If chkLoad = True Then

            Call HeadList()
            Call Search()

        End If

    End Sub

    Private Sub cboTransport_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTransport.SelectedIndexChanged
        If chkLoad = True Then
            chkTransport.Checked = True
            Call HeadList()
            Call Search()

        End If
    End Sub

    Private Sub Date01_ValueChanged(sender As Object, e As EventArgs) Handles Date01.ValueChanged
        If chkLoad = True Then
            chkDate.Enabled = True
            Call HeadList()
            Call Search()

        End If
    End Sub

    Private Sub chkTransport_CheckedChanged(sender As Object, e As EventArgs) Handles chkTransport.CheckedChanged
        If chkLoad = True Then
            Call HeadList()
            Call Search()
        End If

    End Sub

    Private Sub lsvShowBill_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lsvShowBill.SelectedIndexChanged

    End Sub

    Private Sub cboCLoop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCLoop.SelectedIndexChanged
        If chkLoad = True Then
            'chkTransport.Checked = True
            Call HeadList()
            Call Search()

        End If
    End Sub
End Class

