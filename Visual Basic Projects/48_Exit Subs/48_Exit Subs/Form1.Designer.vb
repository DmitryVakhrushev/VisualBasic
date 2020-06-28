<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btnExitSub = New System.Windows.Forms.Button()
        Me.listNumbers = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'btnExitSub
        '
        Me.btnExitSub.Location = New System.Drawing.Point(22, 22)
        Me.btnExitSub.Name = "btnExitSub"
        Me.btnExitSub.Size = New System.Drawing.Size(240, 27)
        Me.btnExitSub.TabIndex = 0
        Me.btnExitSub.Text = "exit sub demo"
        Me.btnExitSub.UseVisualStyleBackColor = True
        '
        'listNumbers
        '
        Me.listNumbers.FormattingEnabled = True
        Me.listNumbers.Location = New System.Drawing.Point(22, 68)
        Me.listNumbers.Name = "listNumbers"
        Me.listNumbers.Size = New System.Drawing.Size(240, 173)
        Me.listNumbers.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.listNumbers)
        Me.Controls.Add(Me.btnExitSub)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnExitSub As System.Windows.Forms.Button
    Friend WithEvents listNumbers As System.Windows.Forms.ListBox

End Class
