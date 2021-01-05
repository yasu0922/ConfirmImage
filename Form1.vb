Imports OpenCvSharp
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.OpenFileDialog1.Reset()
        If (Me.OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK) Then
            Me.TextBox1.Text = Me.OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.OpenFileDialog1.Reset()
        If (Me.OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK) Then
            Me.TextBox2.Text = Me.OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.OpenFileDialog1.Reset()
        If (Me.OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK) Then
            Me.TextBox3.Text = Me.OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        opencv()
    End Sub

    Private Function runRatioTest(matches As DMatch()(), ratio As Double)
        Dim good As New List(Of DMatch)

        ' 配列よりもListのほうが取り回しが良いので変換しておく
        ' .netのFor Eachは配列の要素を変数に分割してくれないから
        ' 配列のまま扱い、 m(0) = m, m(1) = n　に読み替えて考える
        For Each m In matches
            If m(0).Distance < ratio * m(1).Distance Then
                good.Add(m(0))
            End If
        Next
        Return good
    End Function

    Private Function Transform(qry_h As Integer, qry_w As Integer, qry_kp As KeyPoint(), trn_kp As KeyPoint(), good As List(Of DMatch), ByRef matchesMask As Mat)
        Dim qry_pts As List(Of Point2f) = New List(Of Point2f)
        Dim trn_pts As List(Of Point2f) = New List(Of Point2f)
        Dim H
        Dim dst

        'qry_pts = np.float32([qry_kp(m.queryIdx).pt for m in good]).reshape(-1, 1, 2)
        'trn_pts = np.float32([trn_kp(m.trainIdx).pt for m in good]).reshape(-1, 1, 2)
        'Pythonでは座標はリストで扱うが、.NetではKeyPoint構造体で扱うので
        '単純にKeyPointを格納するリストを作ってそこに詰めてやればreshape(-1, 1, 2)と同じになる
        For Each good_pts In good
            qry_pts.Add(qry_kp(good_pts.QueryIdx).Pt)
            trn_pts.Add(trn_kp(good_pts.TrainIdx).Pt)
        Next

        'H, mask = Cv2.FindHomograpy(qry_pts, trn_pts, Cv2.RANSAC, 5.0)
        H = Cv2.FindHomography(InputArray.Create(qry_pts), InputArray.Create(trn_pts), HomographyMethods.Ransac, 5.0, matchesMask)

        'matchesMask = mask.ravel().tolist()
        'pts = np.float([[0,0],[0.qry_h-1],[qry_w-1,qry_h-1],[qry_w-1,0]]).reshape(-1, 1, 2)
        'ptsもqry_pts等と同じくKeyPoint構造体を格納する配列を用意して値を埋め込んでやればreshape(-1, 1, 2)と同じになる
        Dim pts(4) As Point2f
        pts(0) = New Point2f(0, 0)
        pts(1) = New Point2f(0, qry_h - 1)
        pts(2) = New Point2f(qry_w - 1, qry_h - 1)
        pts(3) = New Point2f(qry_w - 1, 0)

        'dst = cv2.perspectiveTransform(pts, H)
        dst = Cv2.PerspectiveTransform(pts, H)

        'return dst, matchesMask
        'dstはPoint2f型の配列
        Return dst

    End Function

    Private Sub opencv()
        ' 変数は可能な限り型推論に頼ってます
        Dim MIN_MATCH_COUNT As Integer = 10
        Dim matchesMask As Mat = New Mat
        Dim dst2f As Point2f()
        Dim h
        Dim w

        Dim query = Cv2.ImRead(TextBox1.Text)
        h = query.Height
        w = query.Width

        Dim train = Cv2.ImRead(TextBox2.Text)
        Dim train2 = train.Clone()

        Dim aka = AKAZE.Create()

        Dim query_kp As KeyPoint()
        Dim query_des As Mat = New Mat

        ' BC42030は無視
        aka.DetectAndCompute(query, Nothing, query_kp, query_des)

        Dim train_kp As KeyPoint()
        Dim train_des As Mat = New Mat

        ' BC42030は無視
        aka.DetectAndCompute(train, Nothing, train_kp, train_des)

        ' HammingとHamming2は違うので注意。
        ' 今回はKnnを使うのでcrossCheckもFalseで良い
        Dim bf = New BFMatcher(NormTypes.Hamming, False)

        Dim matches = bf.KnnMatch(query_des, train_des, 2)

        Dim good = runRatioTest(matches, 0.7)

        If good.Count > MIN_MATCH_COUNT Then
            dst2f = Transform(h, w, query_kp, train_kp, good, matchesMask)

            'train2 = train.copy()
            'train変数の宣言部に記載

            '[np.int32(dst)]の部分
            'Polylinesで使う頂点座標(dst)はPoint型のリストを更にリストしないと受け取ってくれない変態仕様なので注意
            Dim dst As Point()
            dst = Array.ConvertAll(dst2f, Function(pf As Point2f) New Point(CType(pf.X, Int32), CType(pf.Y, Int32)))
            Dim ListOfListOfPoints As List(Of List(Of Point)) = New List(Of List(Of Point))
            Dim ListOfPoints As List(Of Point) = New List(Of Point)(dst)
            ListOfListOfPoints.Add(ListOfPoints)

            'train2 = cv2.Ploylines(train2, [np.int32(dst)], True, 255, 3, cv2.LINE_AA)
            Cv2.Polylines(train2, ListOfListOfPoints, True, New Scalar(255, 0, 0), 3, LineTypes.AntiAlias)

        Else
            MsgBox("十分なマッチングが得られませんでした　-" & Len(good) & "/" & MIN_MATCH_COUNT)
            Exit Sub
        End If

        Dim img_result As Mat = New Mat
        ' matchesMaskはバイト型に変換が必要
        ' Mat.GetArrayがvb.netだと使えないのでMarshal.Copで代用
        ' Mat.Totalだと１個多く配列が作られるので-1する
        Dim byteMask(matchesMask.Total - 1) As Byte
        Marshal.Copy(matchesMask.Data, byteMask, 0, byteMask.Length)

        Cv2.DrawMatches(query, query_kp, train2, train_kp, good, img_result, Nothing, Nothing, byteMask, DrawMatchesFlags.NotDrawSinglePoints)

        ' 大きすぎるので画像サイズを変更
        Cv2.Resize(InputArray.Create(img_result), OutputArray.Create(img_result), New OpenCvSharp.Size, 0.2, 0.2, InterpolationFlags.Lanczos4)

        Cv2.ImShow("Matche - 1", img_result)
        Cv2.WaitKey(0)
        Cv2.DestroyAllWindows()
        Cv2.WaitKey(1)
    End Sub
End Class
