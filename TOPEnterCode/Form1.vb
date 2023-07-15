Imports System.Drawing.Imaging
Imports System.IO
Imports System.Environment
Public Class Form1
    Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Integer
    Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Integer) As Integer
    Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
    Private Declare Function GetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
    Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
    Private Declare Function BitBlt Lib "GDI32" (ByVal srchDC As Integer, ByVal srcX As Integer, ByVal srcY As Integer, ByVal srcW As Integer, ByVal srcH As Integer, ByVal desthDC As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal op As Integer) As Integer
    Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Integer) As Integer
    Private Declare Function DeleteObject Lib "GDI32" (ByVal hObj As Integer) As Integer

    Const SRCCOPY As Integer = &HCC0020



    Private Background As Bitmap

    Private fw, fh As Integer


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim strFileName As String
            Dim arrCommandLineArgs As String()
            Dim objStreamWriter As StreamWriter
            Dim strText As String

            arrCommandLineArgs = GetCommandLineArgs()
            strFileName = "Code.tif"

            If arrCommandLineArgs.GetUpperBound(0) = 4 Then
                CutPicture(CaptureScreen(), strFileName, arrCommandLineArgs(1), arrCommandLineArgs(2), arrCommandLineArgs(3), arrCommandLineArgs(4))
                '                PictureBox1.ImageLocation = strFileName
                strText = OCR(strFileName)

                objStreamWriter = New StreamWriter("Code.txt", False)

                objStreamWriter.WriteLine(strText)
                objStreamWriter.Close()
                End
            Else
                End
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
            End
        End Try
    End Sub
    Protected Function CaptureScreen() As Bitmap
        Dim hsdc, hmdc As Integer
        Dim hbmp, hbmpold As Integer
        Dim r As Integer
        Dim objBitmapReturn As Bitmap

        hsdc = CreateDC("DISPLAY", "", "", "")
        hmdc = CreateCompatibleDC(hsdc)
        fw = GetDeviceCaps(hsdc, 8)
        fh = GetDeviceCaps(hsdc, 10)
        hbmp = CreateCompatibleBitmap(hsdc, fw, fh)
        hbmpold = SelectObject(hmdc, hbmp)
        r = BitBlt(hmdc, 0, 0, fw, fh, hsdc, 0, 0, 13369376)
        hbmp = SelectObject(hmdc, hbmpold)
        r = DeleteDC(hsdc)
        r = DeleteDC(hmdc)
        objBitmapReturn = Image.FromHbitmap(New IntPtr(hbmp))
        DeleteObject(hbmp)
        Return objBitmapReturn
    End Function
    Private Sub CutPicture(ByVal objBitmap As Bitmap, ByVal strOutputName As String, ByVal intX As Integer, ByVal intY As Integer, ByVal intWidth As Integer, ByVal intHeight As Integer)
        Dim objSourceBitmap As Bitmap
        Dim objTargetBitmap As Bitmap
        Dim objSourceRectangleF As RectangleF
        Dim objTargetRectangleF As RectangleF
        Dim objTargetSizeF As SizeF = New SizeF(intWidth, intHeight)
        Dim objGraphics As Graphics

        objSourceBitmap = objBitmap

        objSourceRectangleF = New RectangleF(intX, intY, objTargetSizeF.Width, objTargetSizeF.Height)
        objTargetRectangleF = New RectangleF(0.0F, 0.0F, objTargetSizeF.Width, objTargetSizeF.Height)

        objTargetBitmap = New Bitmap(objTargetSizeF.Width, objTargetSizeF.Height, PixelFormat.Format32bppArgb)
        objGraphics = Graphics.FromImage(objTargetBitmap)
        objGraphics.DrawImage(objSourceBitmap, objTargetRectangleF, objSourceRectangleF, GraphicsUnit.Pixel)
        objTargetBitmap.Save(strOutputName, ImageFormat.Tiff)
        objTargetBitmap.Save("Code.png", ImageFormat.Png)
        objGraphics.Dispose()
        objTargetBitmap.Dispose()
    End Sub
    Private Function OCR(ByVal strFileName As String) As String
        Dim doc As New MODI.Document
        Dim img As MODI.Image
        Dim strReturn As String = ""

        doc.Create(strFileName)
        doc.OCR(MODI.MiLANGUAGES.miLANG_ENGLISH, True, True)
        For i As Integer = 0 To doc.Images.Count - 1
            img = CType(doc.Images.Item(i), MODI.Image)
            If strReturn = "" Then
                strReturn = img.Layout.Text()
            Else
                strReturn = strReturn & img.Layout.Text()
            End If
        Next
        doc.Close(False)

        'Force garbage collection so image file is closed properly
        GC.Collect()
        GC.WaitForPendingFinalizers()

        Return strReturn
    End Function
    Private Function GetRandomName() As String
        Dim intRnd As Integer
        Dim intCnt As Integer
        Dim strReturn As String = ""
        
        For intCnt = 1 To 8
            Randomize()
            intRnd = System.Math.Round(Rnd() * 25) + 1
            strReturn = strReturn & Chr(intRnd + 64)
        Next
        
        Return (strReturn)
    End Function
End Class
