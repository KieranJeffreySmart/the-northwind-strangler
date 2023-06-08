Create new Database Access project using Northwind 2.0 template
Design Employee table: 
	Remove Attachments column
	Remove unique key on Windows Username
	
Design ProductCategories Table:
	Remove ProductCategoryImage column

Form_frmOrderDetails.Form_BeforeInsert: 
	Add Me.OrderDate = Now
	
Fix dbSeeChanges: Find instances of OpenRecordset and add dbSeeChanges parameter
	Find "OpenRecordset" : , dbSeeChanges
	Find ".Execute" : + dbSeeChanges
	
	AddToMRU:ln110: g_dbApp.Execute sql, dbFailOnError + dbSeeChanges
	RemoveFromMRU:ln30
	Form_frmPurchaseOrderDetails.VendorID_BeforeUpdate:ln50
	Form_sfrmAdmin_DeleteTestData.cmdRemoveTestData_Click:ln40-460
	Form_sfrmAdmin_ResetDates.SetDatesToCurrent:ln140
	Form_frmProductDetail:cmdDeleteProduct_Click:ln120-130
	Report_rptProductCatalog:Report_Open:ln40
	
	
	Modinventory.AllocateInventory:ln30
	
	
	
Form_frmOrderDetails:


Private Sub Form_AfterDelConfirm(Status As Integer)
          Dim prodIdStr As String
          Dim productIdIndex As Long
          Dim prodIdLng As Long
          Dim maxProductIdIndex As Long
          
          maxProductIdIndex = List198.ListCount - 1
          For productIdIndex = 0 To maxProductIdIndex
              prodIdStr = List198.Column(0, productIdIndex)
              prodIdLng = CLng(prodIdStr)
              AllocateInventory prodIdLng
          Next productIdIndex
		  
		  10        Dim prodIdStr As String
20        Dim productIdIndex As Long
30        Dim prodIdLng As Long
40        Dim maxProductIdIndex As Long
          
50        maxProductIdIndex = List198.ListCount - 1
60        For productIdIndex = 0 To maxProductIdIndex
70            AllocateInventory CLng(List198.Column(0, productIdIndex))
80        Next productIdIndex
End Sub

Private Sub Form_Delete(Cancel As Integer)
10        On Error GoTo Err_Handler
          
          Dim rmvdProductID, lngProductID As Long
          Dim productsToAllocate() As Long
          Dim productIdsIndex As Long
          
          productIdsIndex = 1
          
20        If MsgBox(GetString(enumStrings.sDeleteRecord, "order"), vbYesNo Or vbQuestion) = vbYes Then

30            lngProductID = 0
              'Delete of the Order causes a cascading delete of order line items
              'We want to delete each line item so we can reallocate inventory for each product
              'before the cascading delete happens
40            With Me.sfrmOrderLineItems.Form.RecordsetClone
50                If .RecordCount > 0 Then
60                    .MoveFirst
70                    While Not .EOF
80                        lngProductID = !ProductID
                          ReDim Preserve productsToAllocate(1 To productIdsIndex)
                          productsToAllocate(productIdsIndex) = lngProductID
                          productIdsIndex = productIdsIndex + 1
110                       .MoveNext
120                   Wend
130               End If
140           End With

              For Each rmvdProductID In productsToAllocate
                  List198.AddItem Item:=Str(rmvdProductID)
              Next


150           RemoveFromMRU "Orders", Me.OrderID
160       Else
170           Cancel = True
180       End If
          
Exit_Handler:
190       Exit Sub

Err_Handler:
200       clsErrorHandler.HandleError "Form_frmOrderDetails", "Form_Delete"
210       Resume Exit_Handler

10        On Error GoTo Err_Handler
          
11        Dim rmvdProductID, lngProductID As Long
12        Dim productsToAllocate() As Long
13        Dim productIdsIndex As Long
          
14        productIdsIndex = 1
          
20        If MsgBox(GetString(enumStrings.sDeleteRecord, "order"), vbYesNo Or vbQuestion) = vbYes Then

30            lngProductID = 0
              'Delete of the Order causes a cascading delete of order line items
              'We want to delete each line item so we can reallocate inventory for each product
              'before the cascading delete happens
40            With Me.sfrmOrderLineItems.Form.RecordsetClone
50                If .RecordCount > 0 Then
60                    .MoveFirst
70                    While Not .EOF
80                        lngProductID = !ProductID
81                        ReDim Preserve productsToAllocate(1 To productIdsIndex)
82                        productsToAllocate(productIdsIndex) = lngProductID
83                        productIdsIndex = productIdsIndex + 1
110                       .MoveNext
120                   Wend
121                   For Each rmvdProductID In productsToAllocate
122                       List198.AddItem Item:=Str(rmvdProductID)
123                   Next
130               End If
140           End With

150           RemoveFromMRU "Orders", Me.OrderID
160       Else
170           Cancel = True
180       End If
          
Exit_Handler:
190       Exit Sub

Err_Handler:
200       clsErrorHandler.HandleError "Form_frmOrderDetails", "Form_Delete"
210       Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

20        AddToMRU "Orders", Me.OrderID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "Form_frmOrderDetails", "Form_AfterInsert"
50        Resume Exit_Handler
End Sub

	
	
Stop validating AutoNumbers on forms: modValidation.IsBoundToRequiredField - Ignore validation of required fields of type 4 (best I could do)
20        IsBoundToRequiredField = ctl.Parent.RecordsetClone.Fields(ctl.ControlSource).Required AND NOT(ctl.Parent.RecordsetClone.Fields(ctl.ControlSource).Type = 4)

	