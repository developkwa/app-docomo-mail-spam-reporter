Option Explicit On
Imports Selenium
Imports Selenium.ChromeDriver

Module DoCoMoSpamReporter
    Dim driver As New Selenium.ChromeDriver

    Sub main()
        Dim sTitle As String
        Dim iErrorCount As Integer

        On Error GoTo ErrorTrap

        iErrorCount = 0
        With driver
            .Start()
            .Window.Maximize()

            .Get("https://mail.smt.docomo.ne.jp/")
            .FindElementByCss("#main-area > button").Click()
            .Wait(1000)

            sTitle = .Title

            .FindElementByName("authid").Click()

            '自動ログインさせる場合
            '.FindElementByName("authid").SendKeys("ここにユーザー名を記入")
            '.FindElementByName("authpass").SendKeys("ここにパスワードを記入")
            '.FindElementByName("subForm").Click()

            Do While sTitle = .Title
                .Wait(100)
            Loop

            If .FindElementsByCss("#inputDevArea2 > input").Count Then
                .FindElementByCss("#inputDevArea2 > input").Clear()
                '                                                     12345678901234567890
                .FindElementByCss("#inputDevArea2 > input").SendKeys("Docomo Spam reporter")
                .FindElementByName("rotpwd").Click()
            End If

            Do While .FindElementsByName("受信BOX").Count = 0
                .Wait(100)
            Loop

            If .FindElementsByName("迷惑メール").Count = 0 Then
                MessageBox.Show("「迷惑メール」フォルダが見つかりませんでした。" & vbCrLf &
                                "「迷惑メール」というフォルダを作成し、そこに迷惑メールを移動して再度実行してください。", vbOKOnly + vbExclamation, "DoCoMo Spam Reporter")
                End
            End If

            .FindElementByName("迷惑メール").Click()

            .Wait(3000)

            Do While .FindElementsByClass("checkbox").Count > 0
                iErrorCount = 0

                '迷惑メールの1番目を選択
                .FindElementByClass("checkbox").Click()

                'その他
                .FindElementByClass("icon-big-etc").Click()

                '迷惑メール報告
                .FindElementByClass("item-spam").Click()

                '迷惑メール送信者の利用停止などに活用させていただきます｡
                'よろしければOKを押してください｡
                .FindElementByClass("btn-ok").Click()

                '迷惑メール報告のため、以下のメールを送信してもよろしいですか？
                .FindElementByClass("btn-ok").Click()

            Loop

        End With
        Exit Sub

ErrorTrap:
        iErrorCount += 1
        If iErrorCount < 10 Then
            System.Threading.Thread.Sleep(100)
            Resume
        Else
            End
        End If
    End Sub
End Module
