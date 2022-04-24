<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormSkompare
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormSkompare))
        Me.BtnNew = New System.Windows.Forms.Button()
        Me.OpenFDNew = New System.Windows.Forms.OpenFileDialog()
        Me.LblNewFile = New System.Windows.Forms.Label()
        Me.LblNewFileName = New System.Windows.Forms.Label()
        Me.LblNewSheets = New System.Windows.Forms.Label()
        Me.LBoxNewSheets = New System.Windows.Forms.ListBox()
        Me.BtnOld = New System.Windows.Forms.Button()
        Me.LblOldFile = New System.Windows.Forms.Label()
        Me.LblOldSheets = New System.Windows.Forms.Label()
        Me.LblOldFileName = New System.Windows.Forms.Label()
        Me.LBoxOldSheets = New System.Windows.Forms.ListBox()
        Me.BtnComp = New System.Windows.Forms.Button()
        Me.TBoxStats = New System.Windows.Forms.RichTextBox()
        Me.BtnStats = New System.Windows.Forms.Button()
        Me.GBoxCompareStyle = New System.Windows.Forms.GroupBox()
        Me.RBtnStyle5 = New System.Windows.Forms.RadioButton()
        Me.RBtnStyle6 = New System.Windows.Forms.RadioButton()
        Me.RBtnStyle4 = New System.Windows.Forms.RadioButton()
        Me.RBtnStyle3 = New System.Windows.Forms.RadioButton()
        Me.RBtnStyle2 = New System.Windows.Forms.RadioButton()
        Me.RBtnStyle1 = New System.Windows.Forms.RadioButton()
        Me.GBoxStatsDiff = New System.Windows.Forms.GroupBox()
        Me.BtnColor = New System.Windows.Forms.Button()
        Me.TBoxColor = New System.Windows.Forms.TextBox()
        Me.LblStringEnd = New System.Windows.Forms.Label()
        Me.LblStringStart = New System.Windows.Forms.Label()
        Me.LblColor = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TBoxStringEnd = New System.Windows.Forms.TextBox()
        Me.TBoxColSelect = New System.Windows.Forms.TextBox()
        Me.TBoxStringStart = New System.Windows.Forms.TextBox()
        Me.LblStart = New System.Windows.Forms.Label()
        Me.TBoxStart = New System.Windows.Forms.TextBox()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.BtnLang = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.BtnGetStartPoint = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GBoxCompareStyle.SuspendLayout()
        Me.GBoxStatsDiff.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnNew
        '
        resources.ApplyResources(Me.BtnNew, "BtnNew")
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.UseVisualStyleBackColor = True
        '
        'OpenFDNew
        '
        Me.OpenFDNew.FileName = "OpenFileDialogNew"
        resources.ApplyResources(Me.OpenFDNew, "OpenFDNew")
        '
        'LblNewFile
        '
        resources.ApplyResources(Me.LblNewFile, "LblNewFile")
        Me.LblNewFile.Name = "LblNewFile"
        '
        'LblNewFileName
        '
        resources.ApplyResources(Me.LblNewFileName, "LblNewFileName")
        Me.LblNewFileName.Name = "LblNewFileName"
        '
        'LblNewSheets
        '
        resources.ApplyResources(Me.LblNewSheets, "LblNewSheets")
        Me.LblNewSheets.Name = "LblNewSheets"
        '
        'LBoxNewSheets
        '
        resources.ApplyResources(Me.LBoxNewSheets, "LBoxNewSheets")
        Me.LBoxNewSheets.FormattingEnabled = True
        Me.LBoxNewSheets.Items.AddRange(New Object() {resources.GetString("LBoxNewSheets.Items")})
        Me.LBoxNewSheets.Name = "LBoxNewSheets"
        '
        'BtnOld
        '
        resources.ApplyResources(Me.BtnOld, "BtnOld")
        Me.BtnOld.Name = "BtnOld"
        Me.BtnOld.UseVisualStyleBackColor = True
        '
        'LblOldFile
        '
        resources.ApplyResources(Me.LblOldFile, "LblOldFile")
        Me.LblOldFile.Name = "LblOldFile"
        '
        'LblOldSheets
        '
        resources.ApplyResources(Me.LblOldSheets, "LblOldSheets")
        Me.LblOldSheets.Name = "LblOldSheets"
        '
        'LblOldFileName
        '
        resources.ApplyResources(Me.LblOldFileName, "LblOldFileName")
        Me.LblOldFileName.Name = "LblOldFileName"
        '
        'LBoxOldSheets
        '
        resources.ApplyResources(Me.LBoxOldSheets, "LBoxOldSheets")
        Me.LBoxOldSheets.FormattingEnabled = True
        Me.LBoxOldSheets.Items.AddRange(New Object() {resources.GetString("LBoxOldSheets.Items")})
        Me.LBoxOldSheets.Name = "LBoxOldSheets"
        '
        'BtnComp
        '
        resources.ApplyResources(Me.BtnComp, "BtnComp")
        Me.BtnComp.Name = "BtnComp"
        Me.BtnComp.UseVisualStyleBackColor = True
        '
        'TBoxStats
        '
        resources.ApplyResources(Me.TBoxStats, "TBoxStats")
        Me.TBoxStats.Name = "TBoxStats"
        '
        'BtnStats
        '
        resources.ApplyResources(Me.BtnStats, "BtnStats")
        Me.BtnStats.Name = "BtnStats"
        Me.BtnStats.UseVisualStyleBackColor = True
        '
        'GBoxCompareStyle
        '
        resources.ApplyResources(Me.GBoxCompareStyle, "GBoxCompareStyle")
        Me.GBoxCompareStyle.Controls.Add(Me.RBtnStyle5)
        Me.GBoxCompareStyle.Controls.Add(Me.RBtnStyle6)
        Me.GBoxCompareStyle.Controls.Add(Me.RBtnStyle4)
        Me.GBoxCompareStyle.Controls.Add(Me.RBtnStyle3)
        Me.GBoxCompareStyle.Controls.Add(Me.RBtnStyle2)
        Me.GBoxCompareStyle.Controls.Add(Me.RBtnStyle1)
        Me.GBoxCompareStyle.Name = "GBoxCompareStyle"
        Me.GBoxCompareStyle.TabStop = False
        '
        'RBtnStyle5
        '
        resources.ApplyResources(Me.RBtnStyle5, "RBtnStyle5")
        Me.RBtnStyle5.Name = "RBtnStyle5"
        Me.RBtnStyle5.UseVisualStyleBackColor = True
        '
        'RBtnStyle6
        '
        resources.ApplyResources(Me.RBtnStyle6, "RBtnStyle6")
        Me.RBtnStyle6.Name = "RBtnStyle6"
        Me.RBtnStyle6.UseVisualStyleBackColor = True
        '
        'RBtnStyle4
        '
        resources.ApplyResources(Me.RBtnStyle4, "RBtnStyle4")
        Me.RBtnStyle4.Name = "RBtnStyle4"
        Me.RBtnStyle4.UseVisualStyleBackColor = True
        '
        'RBtnStyle3
        '
        resources.ApplyResources(Me.RBtnStyle3, "RBtnStyle3")
        Me.RBtnStyle3.Name = "RBtnStyle3"
        Me.RBtnStyle3.UseVisualStyleBackColor = True
        '
        'RBtnStyle2
        '
        resources.ApplyResources(Me.RBtnStyle2, "RBtnStyle2")
        Me.RBtnStyle2.Name = "RBtnStyle2"
        Me.RBtnStyle2.UseVisualStyleBackColor = True
        '
        'RBtnStyle1
        '
        resources.ApplyResources(Me.RBtnStyle1, "RBtnStyle1")
        Me.RBtnStyle1.Name = "RBtnStyle1"
        Me.RBtnStyle1.UseVisualStyleBackColor = True
        '
        'GBoxStatsDiff
        '
        resources.ApplyResources(Me.GBoxStatsDiff, "GBoxStatsDiff")
        Me.GBoxStatsDiff.Controls.Add(Me.BtnColor)
        Me.GBoxStatsDiff.Controls.Add(Me.TBoxColor)
        Me.GBoxStatsDiff.Controls.Add(Me.LblStringEnd)
        Me.GBoxStatsDiff.Controls.Add(Me.LblStringStart)
        Me.GBoxStatsDiff.Controls.Add(Me.LblColor)
        Me.GBoxStatsDiff.Controls.Add(Me.Label1)
        Me.GBoxStatsDiff.Controls.Add(Me.TBoxStringEnd)
        Me.GBoxStatsDiff.Controls.Add(Me.TBoxColSelect)
        Me.GBoxStatsDiff.Controls.Add(Me.TBoxStringStart)
        Me.GBoxStatsDiff.Controls.Add(Me.LblStart)
        Me.GBoxStatsDiff.Controls.Add(Me.TBoxStart)
        Me.GBoxStatsDiff.Name = "GBoxStatsDiff"
        Me.GBoxStatsDiff.TabStop = False
        '
        'BtnColor
        '
        resources.ApplyResources(Me.BtnColor, "BtnColor")
        Me.BtnColor.BackgroundImage = Global.Skompare.My.Resources.Resources.PickIcon
        Me.BtnColor.Name = "BtnColor"
        Me.BtnColor.UseVisualStyleBackColor = True
        '
        'TBoxColor
        '
        resources.ApplyResources(Me.TBoxColor, "TBoxColor")
        Me.TBoxColor.BackColor = System.Drawing.Color.Yellow
        Me.TBoxColor.Name = "TBoxColor"
        '
        'LblStringEnd
        '
        resources.ApplyResources(Me.LblStringEnd, "LblStringEnd")
        Me.LblStringEnd.Name = "LblStringEnd"
        '
        'LblStringStart
        '
        resources.ApplyResources(Me.LblStringStart, "LblStringStart")
        Me.LblStringStart.Name = "LblStringStart"
        '
        'LblColor
        '
        resources.ApplyResources(Me.LblColor, "LblColor")
        Me.LblColor.Name = "LblColor"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'TBoxStringEnd
        '
        resources.ApplyResources(Me.TBoxStringEnd, "TBoxStringEnd")
        Me.TBoxStringEnd.Name = "TBoxStringEnd"
        '
        'TBoxColSelect
        '
        resources.ApplyResources(Me.TBoxColSelect, "TBoxColSelect")
        Me.TBoxColSelect.Name = "TBoxColSelect"
        '
        'TBoxStringStart
        '
        resources.ApplyResources(Me.TBoxStringStart, "TBoxStringStart")
        Me.TBoxStringStart.Name = "TBoxStringStart"
        '
        'LblStart
        '
        resources.ApplyResources(Me.LblStart, "LblStart")
        Me.LblStart.Name = "LblStart"
        '
        'TBoxStart
        '
        resources.ApplyResources(Me.TBoxStart, "TBoxStart")
        Me.TBoxStart.Name = "TBoxStart"
        '
        'BtnLang
        '
        resources.ApplyResources(Me.BtnLang, "BtnLang")
        Me.BtnLang.Name = "BtnLang"
        Me.BtnLang.UseVisualStyleBackColor = True
        '
        'Button1
        '
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.Name = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BtnGetStartPoint
        '
        resources.ApplyResources(Me.BtnGetStartPoint, "BtnGetStartPoint")
        Me.BtnGetStartPoint.Name = "BtnGetStartPoint"
        Me.BtnGetStartPoint.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        resources.ApplyResources(Me.Panel1, "Panel1")
        Me.Panel1.Controls.Add(Me.BtnLang)
        Me.Panel1.Controls.Add(Me.BtnNew)
        Me.Panel1.Controls.Add(Me.LblOldFile)
        Me.Panel1.Controls.Add(Me.LblNewFile)
        Me.Panel1.Controls.Add(Me.LblNewFileName)
        Me.Panel1.Controls.Add(Me.LblOldFileName)
        Me.Panel1.Controls.Add(Me.BtnOld)
        Me.Panel1.Name = "Panel1"
        '
        'FormSkompare
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GBoxCompareStyle)
        Me.Controls.Add(Me.GBoxStatsDiff)
        Me.Controls.Add(Me.TBoxStats)
        Me.Controls.Add(Me.LBoxOldSheets)
        Me.Controls.Add(Me.LBoxNewSheets)
        Me.Controls.Add(Me.LblOldSheets)
        Me.Controls.Add(Me.LblNewSheets)
        Me.Controls.Add(Me.BtnStats)
        Me.Controls.Add(Me.BtnGetStartPoint)
        Me.Controls.Add(Me.BtnComp)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "FormSkompare"
        Me.GBoxCompareStyle.ResumeLayout(False)
        Me.GBoxCompareStyle.PerformLayout()
        Me.GBoxStatsDiff.ResumeLayout(False)
        Me.GBoxStatsDiff.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BtnNew As Button
    Friend WithEvents OpenFDNew As OpenFileDialog
    Friend WithEvents LblNewFile As Label
    Friend WithEvents LblNewFileName As Label
    Friend WithEvents LblNewSheets As Label
    Friend WithEvents LBoxNewSheets As ListBox
    Friend WithEvents BtnOld As Button
    Friend WithEvents LblOldFile As Label
    Friend WithEvents LblOldSheets As Label
    Friend WithEvents LblOldFileName As Label
    Friend WithEvents LBoxOldSheets As ListBox
    Friend WithEvents BtnComp As Button
    Friend WithEvents TBoxStats As RichTextBox
    Friend WithEvents BtnStats As Button
    Friend WithEvents GBoxCompareStyle As GroupBox
    Friend WithEvents RBtnStyle3 As RadioButton
    Friend WithEvents RBtnStyle2 As RadioButton
    Friend WithEvents RBtnStyle1 As RadioButton
    Friend WithEvents RBtnStyle4 As RadioButton
    Friend WithEvents RBtnStyle5 As RadioButton
    Friend WithEvents GBoxStatsDiff As GroupBox
    Friend WithEvents LblStart As Label
    Friend WithEvents TBoxStart As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TBoxColSelect As TextBox
    Friend WithEvents LblStringEnd As Label
    Friend WithEvents LblStringStart As Label
    Friend WithEvents LblColor As Label
    Friend WithEvents TBoxStringEnd As TextBox
    Friend WithEvents TBoxStringStart As TextBox
    Friend WithEvents TBoxColor As TextBox
    Friend WithEvents ColorDialog1 As ColorDialog
    Friend WithEvents BtnColor As Button
    Friend WithEvents RBtnStyle6 As RadioButton
    Friend WithEvents BtnLang As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents BtnGetStartPoint As Button
    Friend WithEvents Panel1 As Panel
End Class
