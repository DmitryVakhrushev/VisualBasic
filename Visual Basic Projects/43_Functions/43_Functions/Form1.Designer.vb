﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.btnRunFunction = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnRunFunction
        '
        Me.btnRunFunction.Location = New System.Drawing.Point(83, 44)
        Me.btnRunFunction.Name = "btnRunFunction"
        Me.btnRunFunction.Size = New System.Drawing.Size(116, 37)
        Me.btnRunFunction.TabIndex = 0
        Me.btnRunFunction.Text = "Run Function"
        Me.btnRunFunction.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 125)
        Me.Controls.Add(Me.btnRunFunction)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnRunFunction As System.Windows.Forms.Button

End Class
