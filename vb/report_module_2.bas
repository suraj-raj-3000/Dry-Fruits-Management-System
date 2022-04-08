Attribute VB_Name = "report_module"
Public Function purchase_report_function()
report.order_report.Visible = False
report.purchase_report.Visible = True
report.sell_report.Visible = False
report.stock_report.Visible = False
report.product_report.Visible = False
report.customer_report.Visible = False
report.supplier_report.Visible = False
report.return_report.Visible = False

order_report_color
purchase_report_color
sale_report_color
stock_report_color
product_report_color
customer_report_color
supplier_report_color
return_report_color
End Function


Public Function return_report_function()
report.order_report.Visible = False
report.purchase_report.Visible = False
report.sell_report.Visible = False
report.stock_report.Visible = False
report.product_report.Visible = False
report.customer_report.Visible = False
report.supplier_report.Visible = False
report.return_report.Visible = True

order_report_color
purchase_report_color
sale_report_color
stock_report_color
product_report_color
customer_report_color
supplier_report_color
return_report_color
End Function

Public Function sale_report_function()
report.order_report.Visible = False
report.purchase_report.Visible = False
report.sell_report.Visible = True
report.stock_report.Visible = False
report.product_report.Visible = False
report.customer_report.Visible = False
report.supplier_report.Visible = False
report.return_report.Visible = False

order_report_color
purchase_report_color
sale_report_color
stock_report_color
product_report_color
customer_report_color
supplier_report_color
return_report_color
End Function

Public Function product_report_function()
report.order_report.Visible = False
report.purchase_report.Visible = False
report.sell_report.Visible = False
report.stock_report.Visible = False
report.product_report.Visible = True
report.customer_report.Visible = False
report.supplier_report.Visible = False
report.return_report.Visible = False

order_report_color
purchase_report_color
sale_report_color
stock_report_color
product_report_color
customer_report_color
supplier_report_color
return_report_color
End Function

Public Function stock_report_function()
report.order_report.Visible = False
report.purchase_report.Visible = False
report.sell_report.Visible = False
report.stock_report.Visible = True
report.product_report.Visible = False
report.customer_report.Visible = False
report.supplier_report.Visible = False
report.return_report.Visible = False

order_report_color
purchase_report_color
sale_report_color
stock_report_color
product_report_color
customer_report_color
supplier_report_color
return_report_color
End Function


Public Function supplier_report_function()
report.order_report.Visible = False
report.purchase_report.Visible = False
report.sell_report.Visible = False
report.stock_report.Visible = False
report.product_report.Visible = False
report.customer_report.Visible = False
report.supplier_report.Visible = True
report.return_report.Visible = False

order_report_color
purchase_report_color
sale_report_color
stock_report_color
product_report_color
customer_report_color
supplier_report_color
return_report_color
End Function

Public Function order_report_function()
report.order_report.Visible = True
report.purchase_report.Visible = False
report.sell_report.Visible = False
report.stock_report.Visible = False
report.product_report.Visible = False
report.customer_report.Visible = False
report.supplier_report.Visible = False
report.return_report.Visible = False

order_report_color
purchase_report_color
sale_report_color
stock_report_color
product_report_color
customer_report_color
supplier_report_color
return_report_color
End Function

Public Function customer_report_function()
report.order_report.Visible = False
report.purchase_report.Visible = False
report.sell_report.Visible = False
report.stock_report.Visible = False
report.product_report.Visible = False
report.customer_report.Visible = True
report.supplier_report.Visible = False
report.return_report.Visible = False

order_report_color
purchase_report_color
sale_report_color
stock_report_color
product_report_color
customer_report_color
supplier_report_color
return_report_color
End Function


Public Function order_report_color()
If report.order_report.Visible = True Then
report.Command1.BackColor = &HFF80FF
Else
report.Command1.BackColor = &H8000000F
End If
End Function

Public Function purchase_report_color()
If report.purchase_report.Visible = True Then
report.Command2.BackColor = &HFF80FF
Else
report.Command2.BackColor = &H8000000F
End If
End Function

Public Function sale_report_color()
If report.sell_report.Visible = True Then
report.Command3.BackColor = &HFF80FF
Else
report.Command3.BackColor = &H8000000F
End If
End Function

Public Function stock_report_color()
If report.stock_report.Visible = True Then
report.Command4.BackColor = &HFF80FF
Else
report.Command4.BackColor = &H8000000F
End If
End Function

Public Function product_report_color()
If report.product_report.Visible = True Then
report.Command5.BackColor = &HFF80FF
Else
report.Command5.BackColor = &H8000000F
End If
End Function

Public Function customer_report_color()
If report.customer_report.Visible = True Then
report.Command6.BackColor = &HFF80FF
Else
report.Command6.BackColor = &H8000000F
End If
End Function

Public Function supplier_report_color()
If report.supplier_report.Visible = True Then
report.Command7.BackColor = &HFF80FF
Else
report.Command7.BackColor = &H8000000F
End If
End Function

Public Function return_report_color()
If report.return_report.Visible = True Then
report.Command8.BackColor = &HFF80FF
Else
report.Command8.BackColor = &H8000000F
End If
End Function
