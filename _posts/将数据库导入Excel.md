title: 将数据库导入EXCEL
---
Option Explicit  '强制申明
Dim cn As New ADODB.Connection
Dim myrs As New ADODB.Recordset
Dim mycnstr As String

Private Sub Command1_Click()
Dim i%, j%
Dim newxls As Excel.application '先声明
Dim newbook As Excel.workbook
Dim newsheet As Excel.worksheet
Set newxls = CreateObject("excel.application") '再定义
Set newbook = newxls.Workbooks.Open("" & App.Path & "\导地线库.xlsx")  '创建工作簿
Set newsheet = newbook.Worksheets(1)
If Adodc1.Recordset.EOF = False Then   '定义了adodc1的数据连接，recordset就已经默认存在了，不需要再打开并使用sql语言
For i = 0 To Adodc1.Recordset.RecordCount - 1
For j = 0 To Adodc1.Recordset.Fields.Count - 1  '遍历数据库的所有记录 记录可以count，字段也有count
On Error Resume Next
DataGrid1.Row = i
DataGrid1.Col = j
newsheet.cells(i + 1, j + 1) = DataGrid1.Text   'datagrid也有行和列，明确了行和列就可以指出具体的值
Next j
Next i
End If

newbook.Close  '记得打开了一定要关闭，先关闭book，再退出程序，最后释放内存
newxls.Quit
Set newxls = Nothing

End Sub


Private Sub Form_Load()
mycnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\jjk.mdb;Persist Security Info=False"
'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\jjk.mdb;Persist Security Info=False"
Adodc1.ConnectionString = mycnstr
Adodc1.CommandType = adCmdTable
Adodc1.RecordSource = "ddxk"    '设置adodc1的数据源
Set DataGrid1.DataSource = Adodc1     '设置datagrid的数据源，把他关联到adodc1上
Text1.Text = App.Path & "\jjk.mbp" '显示数据库的路径
End Sub



