Attribute VB_Name = "函数"
Public Sub ini()
With Form3
    .msf.Clear
    .msf.Rows = 2
    .msf.RowHeightMin = 300
    .msf.Font.Size = 11
    
    .msf.ColAlignment(0) = flexAlignCenterCenter
    .msf.ColAlignment(1) = flexAlignCenterCenter
    .msf.ColAlignment(2) = flexAlignCenterCenter
    .msf.ColAlignment(3) = flexAlignCenterCenter
    .msf.ColAlignment(4) = flexAlignCenterCenter
    .msf.ColAlignment(5) = flexAlignCenterCenter
    .msf.ColAlignment(6) = flexAlignCenterCenter
    .msf.ColAlignment(7) = flexAlignCenterCenter
    
    .msf.TextMatrix(0, 0) = "序号"
    .msf.TextMatrix(0, 1) = "病历编号"
    .msf.TextMatrix(0, 2) = "日期"
    .msf.TextMatrix(0, 3) = "医师"
    .msf.TextMatrix(0, 4) = "治疗类型"
    .msf.TextMatrix(0, 5) = "内容"
    .msf.TextMatrix(0, 6) = "治疗费"
    .msf.TextMatrix(0, 7) = "备注"
    
    
    .msf.ColWidth(0) = 500
    
    .msf.ColWidth(2) = 1400
    
    
    .msf.ColWidth(5) = 6000
    
    .msf.ColWidth(7) = 2000
End With
End Sub
