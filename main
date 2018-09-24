Sub Szukanie_Klientow_v2()
  Const s_InvoiceDataWorksheet As String = "Sheet2"
  Const s_InvoiceDataColumn    As String = "A:A"
  Const s_CustomerWorksheet    As String = "Sheet1"
  Const s_CustomerStartCell    As String = "C2"
  Const s_InvoiceNumPrefix     As String = "418"
  Const n_InvoiceNumLength       As Long = 8
  'Const n_InvScanStartOffset     As Long = -30
  Const n_InvScanEndOffset       As Long = 20
  
  Dim f As Excel.WorksheetFunction: Set f = Excel.WorksheetFunction ' Shortcut

  With Worksheets(s_InvoiceDataWorksheet).Range(s_InvoiceDataColumn)
    With .Parent.Range(.Cells(1), .Cells(Cells.Rows.Count).End(xlUp))
      Dim varInvoiceDataArray As Variant
      varInvoiceDataArray = f.Transpose(.Cells.Value2)
    End With
  End With
  With Worksheets(s_CustomerWorksheet).Range(s_CustomerStartCell)
    With .Parent.Range(.Cells(1), .EntireColumn.Cells(Cells.Rows.Count).End(xlUp))
      Dim varCustomerArray  As Variant
      varCustomerArray = f.Transpose(.Cells.Value2)
    End With
  End With

  Dim varCustomer As Variant
  For Each varCustomer In varCustomerArray
    Dim dblCustomerIndex As Variant
    dblCustomerIndex = Application.Match(varCustomer & "*", varInvoiceDataArray, 0)
    If Not IsError(dblCustomerIndex) _
    And varCustomer <> vbNullString _
    Then
      Dim i As Long
      For i = f.Max(dblCustomerIndex + n_InvScanStartOffset, 1) _
          To f.Min(dblCustomerIndex + n_InvScanEndOffset, UBound(varInvoiceDataArray))
        Dim strInvoiceNum As String
        strInvoiceNum = Right$(Trim$(varInvoiceDataArray(i)), n_InvoiceNumLength)
        If (Left$(strInvoiceNum, Len(s_InvoiceNumPrefix)) = s_InvoiceNumPrefix) Then
          MsgBox "Customer found: " & varCustomer & "." & _
          vbNewLine & _
          "Invoice number: 0" & strInvoiceNum & "*." & _
          vbNewLine & _
          vbNewLine & _
          "*Invoice number can repeat."
        End If
      Next
    End If
  Next varCustomer
  
MsgBox "Searching ended."

End Sub
