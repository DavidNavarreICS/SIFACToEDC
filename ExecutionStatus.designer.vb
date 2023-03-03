<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExecutionStatus
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LblPhase = New System.Windows.Forms.Label()
        Me.ProgressTraitement = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'LblPhase
        '
        Me.LblPhase.AutoSize = True
        Me.LblPhase.Dock = System.Windows.Forms.DockStyle.Top
        Me.LblPhase.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblPhase.Location = New System.Drawing.Point(0, 0)
        Me.LblPhase.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblPhase.Name = "LblPhase"
        Me.LblPhase.Size = New System.Drawing.Size(125, 40)
        Me.LblPhase.TabIndex = 11
        Me.LblPhase.Text = "Label2"
        Me.LblPhase.UseWaitCursor = True
        '
        'ProgressTraitement
        '
        Me.ProgressTraitement.Location = New System.Drawing.Point(18, 80)
        Me.ProgressTraitement.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.ProgressTraitement.Name = "ProgressTraitement"
        Me.ProgressTraitement.Size = New System.Drawing.Size(794, 54)
        Me.ProgressTraitement.TabIndex = 6
        Me.ProgressTraitement.UseWaitCursor = True
        '
        'ExecutionStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(830, 163)
        Me.Controls.Add(Me.LblPhase)
        Me.Controls.Add(Me.ProgressTraitement)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "ExecutionStatus"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Travail en cours"
        Me.TopMost = True
        Me.UseWaitCursor = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LblPhase As Windows.Forms.Label
    Friend WithEvents ProgressTraitement As Windows.Forms.ProgressBar
End Class
