	<!-- Created: 03/09/2007 6:30:04 PM -->
<html>
	<head>
		<meta name="GENERATOR" Content="ASP Express 4.1">
		<title>Brown Data</title>
	

<script language="VB" Runat="server">

Sub MyDataGrid_EditCommand(s As Object, e As DataGridCommandEventArgs )
	MyDataGrid.EditItemIndex = e.Item.ItemIndex
	BindData
End Sub
Sub MyDataGrid_Cancel(Source As Object, e As DataGridCommandEventArgs)
	MyDataGrid.EditItemIndex = -1
	BindData()
End Sub


Sub MyDataGrid_UpdateCommand(s As Object, e As DataGridCommandEventArgs)
	Dim MyConn As OleDbConnection
	Dim MyCommand As OleDbCommand
Dim strConn as string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath("C:\HMS2\HMS\Database\HMS.mdb") & ";"
	Dim txtGuID As textbox = E.Item.cells(0).Controls(0)
	Dim txtName As textbox = E.Item.cells(1).Controls(0)
	Dim txtAddress As textbox = E.Item.cells(2).Controls(0)
	Dim txtCity As textbox = E.Item.cells(3).Controls(0)
	Dim txtState As textbox = E.Item.cells(4).Controls(0)
	Dim txtCompany As textbox = E.Item.cells(5).Controls(0)
	Dim txtDesignation As textbox = E.Item.cells(6).Controls(0)
	Dim txtRoomType As textbox = E.Item.cells(7).Controls(0)
	Dim txtRoomNo As textbox = E.Item.cells(8).Controls(0)
	Dim txtAdvance As textbox = E.Item.cells(9).Controls(0)
	Dim strUpdateStmt As String
	strUpdateStmt =" UPDATE CheckIn_Table SET" & _
	" [GuID] = '" & txtGuID.Text  & "', " & _
	" [Name] = '" & txtName.Text  & "', " & _
	" [Address] = '" & txtAddress.Text  & "', " & _
	" [City] = '" & txtCity.Text  & "', " & _
	" [State] = '" & txtState.Text  & "', " & _
	" [Company] = '" & txtCompany.Text  & "', " & _
	" [Designation] = '" & txtDesignation.Text  & "', " & _
	" [RoomType] = '" & txtRoomType.Text  & "', " & _
	" [RoomNo] = '" & txtRoomNo.Text  & "', " & _
	" [Advance] = '" & txtAdvance.Text  & "'" & _
	" WHERE Label25 = " & e.Item.Cells(-1).Text   

	MyConn = New OleDbConnection(strConn)
	MyCommand = New OleDbCommand(strUpdateStmt, MyConn)
	MyConn.Open()
	MyCommand.ExecuteNonQuery()
	MyDataGrid.EditItemIndex = -1
	BindData
End Sub

Sub Page_Load(Source As Object, E As EventArgs)
		if not Page.IsPostBack then
			BindData
		end if
End Sub

Sub BindData()
'put your databinding code here
If Not Page.IsPostBack Then
	BindData()
End If
End Sub

Sub BindData()
 'put your databinding Code here
 ' Do Not User an Order By Clause when using Sorting!!
End Sub

</script>

<script language="VB" Runat="server">
Sub SortCommand_OnClick(Source As Object, E As DataGridSortCommandEventArgs)
	MySQL = MySQL & " ORDER BY " & E.SortExpression
	BindData()
End Sub
</script></head>
	<body>

<form Runat="server" method="post">
<asp:Datagrid Runat="server"
	Id="MyDataGrid"
	GridLines="Both"
	Border-Width="1"
	cellpadding="0"
	cellspacing="0"
	Headerstyle-Font-Name="Verdana"
	Headerstyle-Font-Size="9"
	Font-Name="Arial"
	Font-Size="9"
	BorderColor="Black"
	OnSortCommand = "SortCommand_OnClick"
	AllowSorting = "True"
	AutogenerateColumns="False"
	OnEditcommand="MyDataGrid_EditCommand"
	OnCancelcommand="MyDataGrid_Cancel"
	OnUpdateCommand="MyDataGrid_UpdateCommand"
>
	<Columns>
		<asp:EditCommandColumn ButtonType="LinkButton" UpdateText="Update" CancelText="Cancel" EditText="Edit" HeaderText="Edit"></asp:EditCommandColumn>
		<asp:BoundColumn DataField="GuID" SortExpression="GuID" HeaderText="GuID"></asp:BoundColumn>
		<asp:BoundColumn DataField="Name" SortExpression="Name" HeaderText="Name"></asp:BoundColumn>
		<asp:BoundColumn DataField="Address" SortExpression="Address" HeaderText="Address"></asp:BoundColumn>
		<asp:BoundColumn DataField="City" SortExpression="City" HeaderText="City"></asp:BoundColumn>
		<asp:BoundColumn DataField="State" SortExpression="State" HeaderText="State"></asp:BoundColumn>
		<asp:BoundColumn DataField="Company" SortExpression="Company" HeaderText="Company"></asp:BoundColumn>
		<asp:BoundColumn DataField="Designation" SortExpression="Designation" HeaderText="Designation"></asp:BoundColumn>
		<asp:BoundColumn DataField="RoomType" SortExpression="RoomType" HeaderText="RoomType"></asp:BoundColumn>
		<asp:BoundColumn DataField="RoomNo" SortExpression="RoomNo" HeaderText="RoomNo"></asp:BoundColumn>
		<asp:BoundColumn DataField="Advance" SortExpression="Advance" HeaderText="Advance"></asp:BoundColumn>
	</Columns>
</asp:DataGrid>

</form>






	</body>
</html>