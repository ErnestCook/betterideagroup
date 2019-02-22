<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStatus
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStatus))
        Me.MyProgressBar = New System.Windows.Forms.ProgressBar
        Me.CurrentRecordNumber = New System.Windows.Forms.Label
        Me.lblDataType = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblMaximumRecord = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lstIssues = New System.Windows.Forms.ListBox
        Me.CalculatedOKCounter = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'MyProgressBar
        '
        Me.MyProgressBar.Location = New System.Drawing.Point(22, 32)
        Me.MyProgressBar.Maximum = 1
        Me.MyProgressBar.Name = "MyProgressBar"
        Me.MyProgressBar.Size = New System.Drawing.Size(311, 47)
        Me.MyProgressBar.Step = 1
        Me.MyProgressBar.TabIndex = 0
        Me.MyProgressBar.Value = 1
        '
        'CurrentRecordNumber
        '
        Me.CurrentRecordNumber.AutoSize = True
        Me.CurrentRecordNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CurrentRecordNumber.Location = New System.Drawing.Point(118, 139)
        Me.CurrentRecordNumber.Name = "CurrentRecordNumber"
        Me.CurrentRecordNumber.Size = New System.Drawing.Size(16, 16)
        Me.CurrentRecordNumber.TabIndex = 1
        Me.CurrentRecordNumber.Text = "0"
        '
        'lblDataType
        '
        Me.lblDataType.AutoSize = True
        Me.lblDataType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDataType.Location = New System.Drawing.Point(33, 102)
        Me.lblDataType.Name = "lblDataType"
        Me.lblDataType.Size = New System.Drawing.Size(76, 16)
        Me.lblDataType.TabIndex = 2
        Me.lblDataType.Text = "Loading..."
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(161, 139)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(21, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "of"
        '
        'lblMaximumRecord
        '
        Me.lblMaximumRecord.AutoSize = True
        Me.lblMaximumRecord.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaximumRecord.Location = New System.Drawing.Point(196, 139)
        Me.lblMaximumRecord.Name = "lblMaximumRecord"
        Me.lblMaximumRecord.Size = New System.Drawing.Size(16, 16)
        Me.lblMaximumRecord.TabIndex = 4
        Me.lblMaximumRecord.Text = "0"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(121, 335)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(94, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lstIssues
        '
        Me.lstIssues.FormattingEnabled = True
        Me.lstIssues.Location = New System.Drawing.Point(36, 226)
        Me.lstIssues.Name = "lstIssues"
        Me.lstIssues.Size = New System.Drawing.Size(270, 95)
        Me.lstIssues.TabIndex = 6
        '
        'CalculatedOKCounter
        '
        Me.CalculatedOKCounter.AutoSize = True
        Me.CalculatedOKCounter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalculatedOKCounter.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.CalculatedOKCounter.Location = New System.Drawing.Point(118, 197)
        Me.CalculatedOKCounter.Name = "CalculatedOKCounter"
        Me.CalculatedOKCounter.Size = New System.Drawing.Size(16, 16)
        Me.CalculatedOKCounter.TabIndex = 7
        Me.CalculatedOKCounter.Text = "0"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(33, 171)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(141, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Daily routes results"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(72, 197)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(28, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "OK"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(120, 327)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(92, 23)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        Me.btnClose.Visible = False
        '
        'frmStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(342, 370)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CalculatedOKCounter)
        Me.Controls.Add(Me.lstIssues)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblMaximumRecord)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblDataType)
        Me.Controls.Add(Me.CurrentRecordNumber)
        Me.Controls.Add(Me.MyProgressBar)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmStatus"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Process Status"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MyProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents CurrentRecordNumber As System.Windows.Forms.Label
    Friend WithEvents lblDataType As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblMaximumRecord As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lstIssues As System.Windows.Forms.ListBox
    Friend WithEvents CalculatedOKCounter As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
