Public Class Form1
    Dim LauncherPath As String
    Dim GamePath As String
    Dim PlayerPath As String

    Public GameList() As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GameList = {"0.6", "0.8", "0.9", "0.95",
            "1.0", "1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.68",
            "1.7", "1.8", "1.9", "2.0", "2.1", "2.2", "2.3", "2.4", "2.5", "2.6",
            "3.0", "3.01", "3.02", "3.05", "3.1", "3.2", "3.3", "3.4"}

        '获取启动目录和游戏数据目录
        LauncherPath = Application.StartupPath
        GamePath = LauncherPath + "\GameData"
        PlayerPath = LauncherPath + "\FlashPlayer\FlashPlayer.exe"

        CheckFile()
    End Sub

    Private Sub CheckFile()
        If Dir(PlayerPath) = "" Then
            MsgBox("找不到游戏运行所需播放器！路径如下：" + PlayerPath, 16, "错误")
            End
        End If

        Dim tempPath As String
        For I As Integer = 0 To GameList.Length - 1
            tempPath = GamePath + "\bvn_" + GameList(I) + "\" + GetLauncherName(I)

            If Dir(tempPath) = "" Then
                MsgBox(tempPath + "文件不存在！")
            Else
                ComboBox1.Items.Add(GameList(I))
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CommandLine As String
        If ComboBox1.SelectedIndex < GetArrayPosition("3.0", GameList) Then
            CommandLine = PlayerPath + " " + GamePath + "\bvn_" + ComboBox1.SelectedItem + "\" + GetLauncherName(ComboBox1.SelectedIndex)
        Else
            CommandLine = GamePath + "\bvn_" + ComboBox1.SelectedItem + "\" + GetLauncherName(ComboBox1.SelectedIndex)
        End If

        'MsgBox(CommandLine)
        If (Shell(CommandLine, 1) = 0) Then
            MsgBox("启动游戏失败！" + PlayerPath, 16, "错误")
        Else
            Me.WindowState = 1
        End If

    End Sub

    '取指定元素的位置，找不到返回-1
    Public Function GetArrayPosition(item As String, array As Array) As Integer
        For I As Integer = 0 To array.Length
            If array(I) = item Then
                Return I
            End If
        Next

        Return -1
    End Function

    Public Function GetLauncherName(item As Integer) As String
        If item < GetArrayPosition("3.0", GameList) Then
            Return "bleachvsnaruto.swf"
        ElseIf item < GetArrayPosition("3.02", GameList) Then
            Return "bleachvsnaruto.exe"
        End If

        Return "launch.exe"
    End Function
End Class
