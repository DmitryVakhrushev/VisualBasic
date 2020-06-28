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
        Me.btnCoercion = New System.Windows.Forms.Button()
        Me.txtN = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnCoercion
        '
        Me.btnCoercion.Location = New System.Drawing.Point(58, 34)
        Me.btnCoercion.Name = "btnCoercion"
        Me.btnCoercion.Size = New System.Drawing.Size(169, 58)
        Me.btnCoercion.TabIndex = 0
        Me.btnCoercion.Text = "Coercion Demo"
        Me.btnCoercion.UseVisualStyleBackColor = True
        '
        'txtN
        '
        Me.txtN.Location = New System.Drawing.Point(58, 122)
        Me.txtN.Name = "txtN"
        Me.txtN.Size = New System.Drawing.Size(169, 20)
        Me.txtN.TabIndex = 1
        Me.txtN.Text = "MyDefaultText"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 207)
        Me.Controls.Add(Me.txtN)
        Me.Controls.Add(Me.btnCoercion)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCoercion As System.Windows.Forms.Button
    Friend WithEvents txtN As System.Windows.Forms.TextBox

End Class
