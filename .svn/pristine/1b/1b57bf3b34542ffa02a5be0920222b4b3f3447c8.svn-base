Imports System.Data
Imports System.Data.SqlClient


Partial Class OA_System_OrgLeft
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Dim MyConnection As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA00030") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00031.aspx")
                Response.End()
            End If

            '********初始化Tree********

            '定義TreeView物件並實體化
            Dim Tree1 As New TreeView
            '定義一個TreeNode並實體化
            Dim tmpNote As New TreeNode
            '設定【根目錄】相關屬性內容
            tmpNote.Text = "首頁"
            tmpNote.Value = "0"
            tmpNote.NavigateUrl = "MOA00031.aspx"
            tmpNote.Target = ""

            'Tree建立該Node
            Tree1.Nodes.Add(tmpNote)

            '********建立樹狀結構********
            '取得根目錄節點
            Dim RootNode As TreeNode
            RootNode = Tree1.Nodes(0)
            Dim rc As String

            '呼叫建立子節點的函數
            rc = AddNodes(RootNode, 0)

            '設定Tree的ImageSet
            Tree1.ImageSet = TreeViewImageSet.Arrows
            '置放於PlaceHolder
            Me.PlaceHolder1.Controls.Add(Tree1)

        End If

    End Sub
    Function AddNodes(ByRef tNode As TreeNode, ByVal PId As String) As String
        '******** 遞迴增加樹結構節點 ********
        Try
            '宣告相關變數
            Dim Da As SqlDataAdapter
            Dim Ds As DataSet
            Dim Dt As DataTable
            Dim SqlTxt As String

            '設定資料來源T-SQL
            SqlTxt = "SELECT * FROM ADMINGROUP"    '請修改您的資料表名稱
            '實體化DataAdapter並且取得資料
            Da = New SqlDataAdapter(SqlTxt, MyConnection)
            '實體化DataSet
            Ds = New DataSet
            '資料填入DataSet
            Da.Fill(Ds)
            '設定DataTable
            Dt = New DataTable
            Dt = Ds.Tables(0)

            MyConnection.Close()

            '定義DataRow承接DataTable篩選的結果
            Dim rows() As DataRow
            '定義篩選的條件
            Dim filterExpr As String
            filterExpr = "PARENT_ORG_UID = '" & PId & "'"
            '資料篩選並把結果傳入Rows
            rows = Dt.Select(filterExpr)

            '如果篩選結果有資料
            If rows.GetUpperBound(0) >= 0 Then

                Dim row As DataRow
                Dim tmpNodeId As String
                Dim tmpsText As String
                Dim NewNode As TreeNode
                Dim rc As String

                '逐筆取出篩選後資料
                For Each row In rows
                    '放入相關變數中
                    tmpNodeId = row(0)
                    tmpsText = row(2)

                    '實體化新節點
                    NewNode = New TreeNode
                    '設定節點各屬性
                    NewNode.Text = tmpsText
                    NewNode.Value = tmpNodeId
                    NewNode.NavigateUrl = "MOA00032.aspx?org_uid=" & tmpNodeId
                    NewNode.Target = "Org_right"

                    '摺疊節點
                    NewNode.Collapse()

                    'NewNode.ImageUrl = "../../image/treeopen.gif"

                    '將節點加入Tree中
                    tNode.ChildNodes.Add(NewNode)

                    '呼叫遞回取得子節點
                    rc = AddNodes(NewNode, tmpNodeId)

                Next
            End If
            '傳回成功訊息
            AddNodes = "Success"

        Catch ex As Exception
            'MsgBox(ex.Message)
            AddNodes = "False"

        End Try
    End Function
End Class
