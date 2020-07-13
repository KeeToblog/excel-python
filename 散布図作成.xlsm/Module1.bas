Attribute VB_Name = "Module1"
Option Explicit
    '��ԍ����p�����ɕϊ�����֐�
    Function ConvertToLetter(iCol As Long) As String
       Dim a As Long
       Dim b As Long
       a = iCol
       ConvertToLetter = ""
       Do While iCol > 0
          a = Int((iCol - 1) / 26)
          b = (iCol - 1) Mod 26
          ConvertToLetter = Chr(b + 65) & ConvertToLetter
          iCol = a
       Loop
    End Function

Sub makechart()
    
    '�ϐ���`
    Dim ohm_cm_min As Long, mA_max As Long, last As Long, i As Long, n As Long
    Dim ws1 As Worksheet
    Dim mA_row As String, ohm_cm_row As String
    Dim title As String
    Dim Legend As Legend
    Dim c As ChartObject

    '�ϐ��ɏ�����ꍞ��
    Set ws1 = Worksheets("Sheet")
    ohm_cm_min = ws1.Range("XFD3").End(xlToLeft).Column
    last = (ohm_cm_min + 1) / 4
    
    'for���[�v�ŃO���t�����o�͂���
    For i = 1 To last
        
        mA_row = ConvertToLetter(4 * i - 3)
        ohm_cm_row = ConvertToLetter(4 * i - 2)
        mA_max = ws1.Range(mA_row & "1048576").End(xlUp).Row

        '�U�z�}���쐬���ăf�[�^�͈͂�I������
        ActiveSheet.Shapes.AddChart2(332, xlXYScatter).Select
'        ActiveChart.SetSourceData Source:=ws1.Range("A3:B102")
        ActiveChart.SetSourceData Source:=ws1.Range(mA_row & "3:" & ohm_cm_row & mA_max)
        ActiveChart.ClearToMatchStyle

        '�U�z�}�̃^�C�g�������
        title = ws1.Range(mA_row & "1").Value
        ActiveChart.HasTitle = True
        ActiveChart.ChartTitle.Characters.Text = title
        
        '�U�z�}�̖}������
        ActiveChart.HasLegend = True
        ActiveChart.SeriesCollection(1).Name = title
        ActiveChart.Legend.Position = xlLegendPositionBottom
    Next

    '�U�z�}�̐���
    n = 1
    ActiveSheet.Range("A1").Select

    For Each c In ActiveSheet.ChartObjects
        c.Top = Cells(105, n).Top
        c.Left = Cells(105, n).Left
        c.Height = 150
        c.Width = 180
        n = n + 4
    Next
    
    '�X�e�b�v17�bws1���G�N�Z���̍őO�ʂɂ����Ă���
    ws1.Activate
   
End Sub
