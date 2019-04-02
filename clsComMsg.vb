Public Class clsComMsg

    '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

    'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
    Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly  vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    BeepPublic Class clsComMsg

        '元はアプリケーションの情報を保持していたが、.NETではアプリケーションの扱いが変わったことと、

        'もともと使用しているのはTitle文字列のみであるため、タイトルを文字列型で保持する
        Private m_sAppTitle As String


    Public WriteOnly Property AppTitle() As String
        Set(ByVal value As String)
            m_sAppTitle = value
        End Set
    End Property


    Public Sub NoticeCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbInformation + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub

    Public Sub NoticeCompletion(Optional ByVal strMsg As String = "")

        Me.NoticeCustom("【" & strMsg & "処理】が正常に終了しました。")

    End Sub

    Public Sub ErrorCustom(ByVal strMsg As String)

        Dim wstrMsg As String
        wstrMsg = vbCrLf & strMsg & vbCrLf
        '    Beep
        'MsgBox(strMsg, vbCritical + vbOKOnly + vbApplicationModal, m_AppObj.Title)
        MsgBox(strMsg, vbCritical + vbOKOnly + vbApplicationModal, m_sAppTitle)

    End Sub
    Public Sub ErrorRequired(ByVal strItemName As String)

        Dim wstrMsg As String
        wstrMsg = "【" & strItemName & "】は省略できません。"
        ErrorCustom(wstrMsg)

    End Sub
    Public Sub ErrorNoMaster(Optional ByVal strMasterName As String = "")

        Dim wstrMsg As String
        wstrMsg = "入力されたコードは、【" & strMasterName & "】に登録されていません。"
        ErrorCustom(wstrMsg)

    End Sub

    Public Sub ErrorNoCondition()

        'ErrorCustom(MSG21)

    End Sub
    Public Sub ErrorCompare()

        ErrorCustom("大小関係に誤りがあります。")

    End Sub

    Public Sub ErrorUpdated()

        ErrorCustom("このデータは他ユーザにより、追加・削除もしくは更新されました。" & vbCrLf & _
                    "入力された内容は保存されませんでした。" & vbCrLf & vbCrLf & _
                    "必要に応じて再度処理をやり直してください。")

    End Sub

    Public Sub ErrorAbNormalCompletion(Optional ByVal strMsg As String = "")

        ErrorCustom("【" & strMsg & "処理】は異常終了しました。")

    End Sub

    Public Sub ErrorTimeOut(Optional ByVal strMsg As String = "")

        Dim wstrMsg As String
        wstrMsg = strMsg & vbCrLf & vbCrLf & _
                  "情報の取得(SELECT)および操作(INSERT/UPDATE/DELETE)の制限時間を超えました。" & vbCrLf & _
                  "他のクライアントで更新中の可能性があります。" & vbCrLf & vbCrLf & _
                  "確認してください。"
        ErrorCustom(wstrMsg)

    End Sub

    Public Sub ErrorSystem(ByVal objErr As ErrObject, Optional ByVal strMsg As String = "")

        Dim wstrMsg As String

        '' エラーが発生していないときは、即復帰する
        If objErr.Number = 0 Then
            Exit Sub
        End If

        '' -2147217864 : BatchUpdateにおける他クライアント更新済エラー
        '' -2147217900 : BatchUpdateにおける他クライアント挿入済エラー
        If objErr.Number = -2147217864 Then
            '        objErr.Number = -2147217900 Then
            Me.ErrorUpdated()
            objErr.Clear()
            Exit Sub
        End If

        '' -2147217871 ：時間切れ

        If objErr.Number = -2147217871 Then
            Me.ErrorTimeOut()
            objErr.Clear()
            Exit Sub
        End If

        '' ログインできないとき

        If objErr.Number = -2147217843 Then
            Me.ErrorCustom("ユーザ：" & strMsg & vbCrLf & _
                            "この利用者ではログインできません。" & vbCrLf & _
                            "データベースに対する処理がされていない可能性もあります。環境を見直してください。" & vbCrLf & _
                            "処理を中止します。")
            objErr.Clear()
            Exit Sub
        End If

        wstrMsg = strMsg & vbCrLf
        wstrMsg = wstrMsg & "【Description】 " & objErr.Description & vbCrLf
        wstrMsg = wstrMsg & "【Number】 " & objErr.Number & vbCrLf
        wstrMsg = wstrMsg & "【Souce】 " & objErr.Source & vbCrLf
        wstrMsg = wstrMsg & vbCrLf & "システム管理者に連絡してください。"

        ErrorCustom(wstrMsg)

        objErr.Clear()

    End Sub

    Public Sub ErrorUnique(ByVal strMsg As String)

        ' strMsg = 項目名

        ErrorCustom("【" & strMsg & "】が重複しています。")

    End Sub

    Public Function ConfirmCustom(Optional ByVal strMsg As String = "処理を実行してもよろしいですか？", Optional ByVal blnDefYes As Boolean = True) As Boolean

        Dim wstrMsg As String
        Dim wstyle As MsgBoxStyle

        ConfirmCustom = False
        wstrMsg = strMsg & vbCrLf
        wstyle = vbQuestion + vbYesNo + vbApplicationModal + vbDefaultButton1
        If Not blnDefYes Then
            wstyle = wstyle + vbDefaultButton2
        End If
        '    Beep
        'If MsgBox(wstrMsg, wstyle, m_AppObj.Title) = vbYes Then
        If MsgBox(wstrMsg, wstyle, m_sAppTitle) = vbYes Then
            ConfirmCustom = True
        End If

    End Function

    Public Function ConfirmDelete(Optional ByVal blnDefYes As Boolean = False) As Boolean

        ConfirmDelete = False
        If Me.ConfirmCustom("表示中のデータを削除してもよろしいですか？", blnDefYes) Then
            ConfirmDelete = True
        End If

    End Function

    Public Function ConfirmLineDelete(Optional ByVal blnDefYes As Boolean = False) As Boolean

        ConfirmLineDelete = False
        If Me.ConfirmCustom("該当行を削除してもよろしいですか？", blnDefYes) Then
            ConfirmLineDelete = True
        End If

    End Function

    Public Function ConfirmRegistration(Optional ByVal blnDefYes As Boolean = True) As Boolean

        ConfirmRegistration = False
        If Me.ConfirmCustom("表示されているデータを登録します。よろしいですか？", blnDefYes) Then
            ConfirmRegistration = True
        End If

    End Function

    Public Function ConfirmExecution(Optional ByVal strMsg As String = "", Optional ByVal blnDefYes As Boolean = True) As Boolean

        ConfirmExecution = False
        If Me.ConfirmCustom(strMsg & "処理を実行します。よろしいですか？", blnDefYes) Then
            ConfirmExecution = True
        End If

    End Function

    Public Function ConfirmCancellation(Optional ByVal strMsg As String = "", Optional ByVal blnDefYes As Boolean = True) As Boolean

        ConfirmCancellation = False
        If Me.ConfirmCustom("編集内容を取消してよろしいですか？", blnDefYes) Then
            ConfirmCancellation = True
        End If

    End Function

    Public Function ConfirmEnd(Optional ByVal strMsg As String = "", Optional ByVal blnDefYes As Boolean = False) As Boolean

        ConfirmEnd = False
        If Me.ConfirmCustom("処理を終了します。" & vbCrLf & "編集中の時はそのデータは取り消されます。よろしいですか？", blnDefYes) Then
            ConfirmEnd = True
        End If

    End Function



End Class