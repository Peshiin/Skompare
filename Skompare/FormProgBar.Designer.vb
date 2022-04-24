<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormProgBar
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormProgBar))
        Me.ProgBar = New System.Windows.Forms.ProgressBar()
        Me.LblProgBar = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ProgBar
        '
        resources.ApplyResources(Me.ProgBar, "ProgBar")
        Me.ProgBar.Name = "ProgBar"
        '
        'LblProgBar
        '
        resources.ApplyResources(Me.LblProgBar, "LblProgBar")
        Me.LblProgBar.Name = "LblProgBar"
        '
        'FormProgBar
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.LblProgBar)
        Me.Controls.Add(Me.ProgBar)
        Me.MaximizeBox = False
        Me.Name = "FormProgBar"
        Me.ShowIcon = False
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ProgBar As ProgressBar
    Friend WithEvents LblProgBar As Label
End Class
