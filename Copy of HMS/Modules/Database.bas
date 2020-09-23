Attribute VB_Name = "Database"
Public cnn As ADODB.Connection
Public Rs_Rate As ADODB.Recordset
Public RS_Password As ADODB.Recordset
Public rs As ADODB.Recordset
Public RS_Guest As ADODB.Recordset
Public RS_GuestIn As ADODB.Recordset
Public RS_GuestOut As ADODB.Recordset
Public RS_Enter As ADODB.Recordset
Public RS_Delete As ADODB.Recordset
Public RS_SingleRoom As ADODB.Recordset
Public RS_DoubleRoom As ADODB.Recordset
Public RS_SuiteRoom As ADODB.Recordset
Public RS_DeluxeSuite As ADODB.Recordset
Public Rs_Details As ADODB.Recordset
Public Rs_Detail As ADODB.Recordset
Public RS_Edit As ADODB.Recordset
Public RS_Company As ADODB.Recordset
Public RS_Userlog As ADODB.Recordset
Public RS_Payment As ADODB.Recordset
Public RS_Payroll As ADODB.Recordset
Public RS_DrinkPrice As ADODB.Recordset
Public RS_Paymentlog As ADODB.Recordset
Public RS_Drink As ADODB.Recordset
Public RS_Reservation As ADODB.Recordset
Public RS_Payrolllog As ADODB.Recordset
Public RS_ticker As ADODB.Recordset

Public UserName As String
Public Rights As String
Public Company As String
Public Add As String

Sub Connect()
Set cnn = New ADODB.Connection
Set Rs_Rate = New ADODB.Recordset
Set RS_Password = New ADODB.Recordset
Set rs = New ADODB.Recordset
Set RS_Guest = New ADODB.Recordset
Set RS_SingleRoom = New ADODB.Recordset
Set RS_DoubleRoom = New ADODB.Recordset
Set RS_SuiteRoom = New ADODB.Recordset
Set RS_DeluxeSuite = New ADODB.Recordset
Set RS_GuestIn = New ADODB.Recordset
Set RS_GuestOut = New ADODB.Recordset
Set RS_Enter = New ADODB.Recordset
Set RS_Delete = New ADODB.Recordset
Set Rs_Details = New ADODB.Recordset
Set Rs_Detail = New ADODB.Recordset
Set RS_Edit = New ADODB.Recordset
Set RS_Company = New ADODB.Recordset
Set RS_Userlog = New ADODB.Recordset
Set RS_Payment = New ADODB.Recordset
Set RS_Payroll = New ADODB.Recordset
Set RS_DrinkPrice = New ADODB.Recordset
Set RS_Paymentlog = New ADODB.Recordset
Set RS_Drink = New ADODB.Recordset
Set RS_Reservation = New ADODB.Recordset
Set RS_Payrolllog = New ADODB.Recordset
Set RS_ticker = New ADODB.Recordset

cnn.CursorLocation = adUseClient
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
         & App.Path & ".\Database\HMS.mdb;Persist Security Info=False"
         
Rs_Rate.Open "SELECT * FROM Rate_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Password.Open "SELECT * FROM Password_Table", cnn, adOpenDynamic, adLockOptimistic
rs.Open "SELECT * FROM PresentEmp_Table", cnn, adOpenDynamic, adLockOptimistic
RS_Guest.Open "SELECT * FROM CheckIn_Table", cnn, adOpenDynamic, adLockOptimistic
RS_SingleRoom.Open "SELECT * FROM SingleRoom_Table", cnn, adOpenDynamic, adLockPessimistic
RS_DoubleRoom.Open "SELECT * FROM DoubleRoom_Table", cnn, adOpenDynamic, adLockPessimistic
RS_SuiteRoom.Open "SELECT * FROM SuiteRoom_Table", cnn, adOpenDynamic, adLockPessimistic
RS_DeluxeSuite.Open "SELECT * FROM DeluxeSuite_Table", cnn, adOpenDynamic, adLockPessimistic
RS_GuestIn.Open "SELECT * FROM CheckIn_Table", cnn, adOpenDynamic, adLockPessimistic
RS_GuestOut.Open "SELECT * FROM CheckOut_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Enter.Open "SELECT * FROM DeletedEmp_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Delete.Open "SELECT * FROM PresentEmp_Table", cnn, adOpenDynamic, adLockOptimistic
Rs_Details.Open "SELECT * FROM PresentEmp_Table", cnn, adOpenDynamic, adLockPessimistic
Rs_Detail.Open "SELECT * FROM CheckIn_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Edit.Open "SELECT * FROM CheckIn_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Company.Open "SELECT * FROM Company_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Userlog.Open "SELECT * FROM userlog_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Payment.Open "SELECT * FROM Payment_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Payroll.Open "SELECT * FROM Payroll_Table", cnn, adOpenDynamic, adLockPessimistic
RS_DrinkPrice.Open "SELECT * FROM Drink_price", cnn, adOpenDynamic, adLockPessimistic
RS_Paymentlog.Open "SELECT * FROM Payment_log", cnn, adOpenDynamic, adLockPessimistic
RS_Drink.Open "SELECT * FROM Drinks", cnn, adOpenDynamic, adLockPessimistic
RS_Reservation.Open "SELECT * FROM Reservation_Table", cnn, adOpenDynamic, adLockPessimistic
RS_Payrolllog.Open "SELECT * FROM Payroll_log", cnn, adOpenDynamic, adLockPessimistic
RS_ticker.Open "SELECT * FROM Ticker", cnn, adOpenDynamic, adLockPessimistic


With RS_Company
Company = .Fields(0)
Add = .Fields(1)
End With
End Sub



