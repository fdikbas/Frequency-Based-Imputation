'*************************************************************************************************
'*******************      F R E Q U E N C Y   B A S E D   I M P U T A T I O N        *************
'*******************             ( V I S U A L   B A S I C   C O D E )               *************
'*************************************************************************************************
'
' !!!   IMPORTANT NOTE   !!!
' THIS SOFTWARE IS PROVIDED FOR BEING USED WITH ACADEMIC OR TESTING PURPOSES.
' IF YOU WANT TO USE THE CODE WITH COMMERCIAL PURPOSES PLEASE CONTACT THE AUTHOR AT:
' f_dikbas@pau.edu.tr  OR  fdikbas@yahoo.com
'
' Copyright (c) 2018 Fatih DIKBAS 
' e-mail:f_dikbas@pau.edu.tr
' Address:
' Pamukkale Universitesi, Insaat Muhendisligi Bolumu,
' Kinikli Kampusu, Denizli, Turkey
'
' This program Is free software: you can redistribute it And/Or modify it under the terms Of the GNU General Public License As published by the Free Software Foundation, either version 3 Of the License, Or (at your option) any later version.
' 
' This program Is distributed In the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty Of MERCHANTABILITY Or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License For more details.
' 
' To receive a copy Of the GNU General Public License see: <http://www.gnu.org/licenses/>.
'
' This file is provided as an Appendix of the article titled ‘Compositional Correlation for Detecting Real Associations Among Time Series’ which was submitted to the international publishing company Gece Kitapligi for possible publication.
'
'
' !!!   AN IMPORTANT NOTE   !!!
' Do not click on the cells of the Excel files being generated when the program is making calculations. 
' Otherwise the program will stop making calculations (This is not a problem related with coding).
' But you can browse through and zoom in and out by using sidebars and mouse wheel.

Imports Microsoft.Office.Interop.Excel
'For enabling interoperability with Excel please select: 
'Project - Add Reference - Microsoft.Office.Interop.Excel 
Module Module1
    Sub Main()
        ' Create new Application.
        Dim excel As Application = New Application

        ' Open Excel spreadsheet.
        Dim w As Workbook = excel.Workbooks.Open("C:\VB\FrequencyBasedImputation\Sample.Input.Data.xlsx")
        w.Application.Visible = True
        w.SaveAs("C:\VB\FrequencyBasedImputation\Sample.Outputs.09.02.2015.xlsx")

        Dim w2 As Workbook = excel.Workbooks.Add
        w2.SaveAs("C:\VB\FrequencyBasedImputation\Sample.Comparisons.09.02.2015.xlsx")
        w2.Application.Visible = True

        Dim w3 As Workbook = excel.Workbooks.Add
        w3.SaveAs("C:\VB\FrequencyBasedImputation\Sample.Correlation.Tables.09.02.2015.xlsx")
        w3.Application.Visible = True

        Dim w4 As Workbook = excel.Workbooks.Add
        w4.SaveAs("C:\VB\FrequencyBasedImputation\Sample.Cumulative.Statistics.09.02.2015.xlsx")
        w4.Application.Visible = True

        Dim i, j, k, m As Integer

        ' Get sheet.
        Dim sheet As Worksheet = w.Sheets(1)
        Dim sheet2 As Worksheet = w2.Sheets(1)
        Dim sheet3 As Worksheet = w3.Sheets(1)
        Dim sheet4 As Worksheet = w4.Sheets(1)

        Dim DataFileRange As Range = sheet.UsedRange

        'Load all cells into 2d array:
        Dim ObservedData(,) As Object = DataFileRange.Value(XlRangeValueDataType.xlRangeValueDefault)

        'Scan the cells:
        If ObservedData IsNot Nothing Then Console.WriteLine("Number of Observations in the Data File: {0}", ObservedData.Length)

        'Get bounds of the array (First column includes the text headings for the rows and
        'the first row includes the text headings of the columns)
        Dim NumberOfRows As Integer = ObservedData.GetUpperBound(0)
        Dim NumberOfColumns As Integer = ObservedData.GetUpperBound(1)

        Console.WriteLine("Number of Rows    = {0}", NumberOfRows)
        Console.WriteLine("Number of Columns = {0}", NumberOfColumns)

        Dim rng, rng3 As Range
        Dim rng2 As Range
        Dim cs As ColorScale

        rng2 = sheet.Range(sheet.Cells(NumberOfRows + 6, NumberOfColumns + 10), sheet.Cells(NumberOfRows * 2 + 10, NumberOfColumns * 2 + 14))  ' For the table showing cluster numbers on the right
        rng3 = sheet.Range(sheet.Cells(NumberOfRows + 6, 2), sheet.Cells(NumberOfRows * 2 + 10, NumberOfColumns + 6))
        cs = rng2.FormatConditions.AddColorScale(ColorScaleType:=3)
        cs = rng3.FormatConditions.AddColorScale(ColorScaleType:=3)

        ' Set the color of the lowest value, with a range up to
        ' the next scale criteria. The color should be red.
        With cs.ColorScaleCriteria(1).Type = XlConditionValueTypes.xlConditionValueHighestValue
            With cs.ColorScaleCriteria(1).FormatColor
                .Color = &H6B69F8
                .TintAndShade = 0
            End With
        End With

        ' At the 50th percentile, the color should be red/green.
        ' Note that you cannot set the Value property for all
        ' values of Type.
        With cs.ColorScaleCriteria(2).Type = XlConditionValueTypes.xlConditionValuePercentile
            cs.ColorScaleCriteria(2).Value = 50
            With cs.ColorScaleCriteria(2).FormatColor
                .Color = &H84EBFF
                .TintAndShade = 0
            End With
        End With

        ' At the highest value, the color should be green.
        With cs.ColorScaleCriteria(3).Type = XlConditionValueTypes.xlConditionValueHighestValue
            With cs.ColorScaleCriteria(3).FormatColor
                .Color = &H7BBE63
                .TintAndShade = 0
            End With
        End With

        ' Generation of the expanded data table to be completed
        sheet.Range(sheet.Cells(NumberOfRows + 9, 1), sheet.Cells(NumberOfRows * 2 + 7, 1)).Value2 = sheet.Range(sheet.Cells(2, 1), sheet.Cells(NumberOfRows, 1)).Value2  ' Years
        sheet.Range(sheet.Cells(NumberOfRows + 5, 2), sheet.Cells(NumberOfRows + 5, 4)).Value2 = sheet.Range(sheet.Cells(1, NumberOfColumns - 2), sheet.Cells(1, NumberOfColumns)).Value2   ' August-September txt
        sheet.Range(sheet.Cells(NumberOfRows + 10, 2), sheet.Cells(NumberOfRows * 2 + 8, 4)).NumberFormat = "####.###"
        sheet.Range(sheet.Cells(NumberOfRows + 10, 2), sheet.Cells(NumberOfRows * 2 + 8, 4)).Value2 = sheet.Range(sheet.Cells(2, NumberOfColumns - 2), sheet.Cells(NumberOfRows, NumberOfColumns)).Value2  ' August-September Data
        sheet.Range(sheet.Cells(NumberOfRows + 5, 5), sheet.Cells(NumberOfRows + 5, NumberOfColumns + 3)).Value2 = sheet.Range(sheet.Cells(1, 2), sheet.Cells(1, NumberOfColumns)).Value2   ' October-September txt
        sheet.Range(sheet.Cells(NumberOfRows + 9, 5), sheet.Cells(NumberOfRows * 2 + 7, NumberOfColumns + 3)).NumberFormat = "####.###"
        sheet.Range(sheet.Cells(NumberOfRows + 9, 5), sheet.Cells(NumberOfRows * 2 + 7, NumberOfColumns + 3)).Value2 = sheet.Range(sheet.Cells(2, 2), sheet.Cells(NumberOfRows, NumberOfColumns)).Value2  ' October-September Data
        sheet.Range(sheet.Cells(NumberOfRows + 5, NumberOfColumns + 4), sheet.Cells(NumberOfRows + 5, NumberOfColumns + 6)).Value2 = sheet.Range(sheet.Cells(1, 2), sheet.Cells(1, 4)).Value2   ' October-November txt
        sheet.Range(sheet.Cells(NumberOfRows + 8, NumberOfColumns + 4), sheet.Cells(NumberOfRows * 2 + 6, NumberOfColumns + 6)).NumberFormat = "####.###"
        sheet.Range(sheet.Cells(NumberOfRows + 8, NumberOfColumns + 4), sheet.Cells(NumberOfRows * 2 + 6, NumberOfColumns + 6)).Value2 = sheet.Range(sheet.Cells(2, 2), sheet.Cells(NumberOfRows, 4)).Value2  ' October-November Data

        'Coloring
        sheet.Range(sheet.Cells(NumberOfRows + 5, 2), sheet.Cells(NumberOfRows + 5, NumberOfColumns + 6)).Interior.Color = RGB(210, 225, 245)                        'Light Blue
        sheet.Range(sheet.Cells(NumberOfRows + 9, 1), sheet.Cells(NumberOfRows * 2 + 7, 1)).Interior.Color = RGB(210, 225, 245)                                      'Light Blue
        sheet.Range(sheet.Cells(NumberOfRows + 6, NumberOfColumns + 7), sheet.Cells(NumberOfRows * 2 + 10, NumberOfColumns + 7)).Interior.Color = RGB(160, 190, 230) 'Dark Blue
        sheet.Range(sheet.Cells(NumberOfRows + 4, 2), sheet.Cells(NumberOfRows + 4, NumberOfColumns + 6)).Interior.Color = RGB(160, 190, 230)                        'Dark Blue
        sheet.Range(sheet.Cells(NumberOfRows + 6, 2), sheet.Cells(NumberOfRows * 2 + 10, NumberOfColumns + 6)).Interior.Color = RGB(240, 240, 240)                   'Gray

        'Adding row and column titles to Comparisons file, writing the observed data and coloring
        sheet2.Range(sheet2.Cells(1, 1), sheet2.Cells(NumberOfRows * 6 + 1, NumberOfColumns + 8)).Interior.Color = RGB(240, 240, 240)   'Gray
        sheet2.Cells(1, NumberOfColumns + 2).Interior.Color = RGB(190, 140, 150)                                                        'Dark Lilac  
        sheet2.Cells(1, NumberOfColumns + 3).Interior.Color = RGB(200, 150, 160)                                                        'Darker Lilac
        sheet2.Range(sheet2.Cells(2, 3), sheet2.Cells(NumberOfRows * 6 + 1, NumberOfColumns + 8)).NumberFormat = "####.###"
        sheet2.Range(sheet2.Cells(1, 1), sheet2.Cells(NumberOfRows * 6 + 1, NumberOfColumns + 8)).Font.Bold = True
        With sheet2.Range(sheet2.Cells(1, 3), sheet2.Cells(1, NumberOfColumns + 1))
            .Value2 = sheet.Range(sheet.Cells(1, 2), sheet.Cells(1, NumberOfColumns)).Value2   'Header Row
            .Interior.Color = RGB(160, 190, 230)                                               'Dark Blue
            .Font.Bold = True
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
        End With

        'Adding row and column titles to Correlation Tables file, writing the observed data and coloring
        sheet3.Range(sheet3.Cells(1, 1), sheet3.Cells(NumberOfRows * 8 + 1, NumberOfColumns + 2)).Interior.Color = RGB(240, 240, 240)   'Gray
        sheet3.Range(sheet3.Cells(2, 3), sheet3.Cells(NumberOfRows * 8 + 1, NumberOfColumns * 2 + 8)).NumberFormat = "####.###"
        sheet3.Range(sheet3.Cells(1, 1), sheet3.Cells(NumberOfRows * 8 + 1, NumberOfColumns + 2)).Font.Bold = True
        With sheet3.Range(sheet3.Cells(1, 3), sheet3.Cells(1, NumberOfColumns + 1))
            .Value2 = sheet.Range(sheet.Cells(1, 2), sheet.Cells(1, NumberOfColumns)).Value2   'Header Row
            .Interior.Color = RGB(160, 190, 230)                                               'Dark Blue
            .Font.Bold = True
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
        End With

        'Cumulative Statistics and Graphs dosyasına satır sütun başlıklarının ve gözlenmiş verilerin yerleştirilmesi ve renklendirme:
        'Adding row and column titles to Cumulative Statistics and Graphs file, writing the observed data and coloring
        sheet4.Range(sheet4.Cells(1, 1), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 25)).Interior.Color = RGB(240, 240, 240)   'Gray
        sheet4.Range(sheet4.Cells(2, 2), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 3)).NumberFormat = "####.###"

        sheet4.Columns("A").Font.Bold = True
        With sheet4.Range(sheet4.Cells(1, 2), sheet4.Cells(1, 6))
            .Font.Bold = True
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
        End With

        sheet4.Cells(1, 2).Value2 = "Observed"
        sheet4.Cells(1, 2).Interior.Color = RGB(235, 200, 200)                                                'Dark Pink
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 2).Interior.Color = RGB(225, 190, 190)   'Lighter Dark Pink
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 2).Interior.Color = RGB(235, 200, 200)   'Dark Pink
        sheet4.Cells(1, 3).Value2 = "Estimated"
        sheet4.Cells(1, 3).Interior.Color = RGB(242, 220, 220)                                                'Light Pink
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 3).Interior.Color = RGB(232, 210, 210)   'Lighter Pink
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 3).Interior.Color = RGB(242, 220, 220)   'Light Pink
        sheet4.Cells(1, 4).Value2 = "(E-O)^2"
        sheet4.Cells(1, 4).Interior.Color = RGB(253, 233, 217)                                                'Light Orange
        sheet4.Range(sheet4.Cells(2, 4), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 4)).Interior.Color = RGB(228, 223, 236)   'Light Lilac
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 4).Interior.Color = RGB(243, 223, 207)   'Lighter Orange
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 4).Interior.Color = RGB(253, 233, 217)   'Light Orange
        sheet4.Cells(1, 5).Value2 = "(O-AV)^2"
        sheet4.Cells(1, 5).Interior.Color = RGB(252, 213, 180)                                                'Orange
        sheet4.Range(sheet4.Cells(2, 5), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 5)).Interior.Color = RGB(204, 192, 218)   'Lilac
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 5).Interior.Color = RGB(242, 203, 170)   'Lighter Orange
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 5).Interior.Color = RGB(252, 213, 180)   'Orange
        sheet4.Cells(1, 6).Value2 = "ABS(O-E)"
        sheet4.Cells(1, 6).Interior.Color = RGB(250, 191, 143)                                                'Dark Orange
        sheet4.Range(sheet4.Cells(2, 6), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 6)).Interior.Color = RGB(177, 160, 200)   'Dark Lilac
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 6).Interior.Color = RGB(240, 181, 133)                'Lighter Dark Orange
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 6).Interior.Color = RGB(250, 191, 143)                'Dark Orange
        sheet4.Cells(1, 7).Value2 = "Est.(1-2)"
        sheet4.Cells(1, 8).Value2 = "Est.(1-3)"
        sheet4.Cells(1, 9).Value2 = "Est.(1-4)"
        sheet4.Cells(1, 10).Value2 = "Est.(1-5)"

        sheet4.Range(sheet4.Cells(2, 2), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 10, 10)).NumberFormat = "0.000"

        For j = 0 To 4
            sheet2.Cells(1, NumberOfColumns + 4 + j).Interior.Color = RGB(220 - j * 10, 230 - j * 10, 190 - j * 10)        'Green Tones                                                                                 'Green Tones
        Next

        sheet2.Range(sheet2.Cells(1, NumberOfColumns + 2), sheet2.Cells(NumberOfRows * 6 + 1, NumberOfColumns + 8)).HorizontalAlignment = XlHAlign.xlHAlignCenter
        sheet3.Range(sheet3.Cells(1, NumberOfColumns + 2), sheet3.Cells(NumberOfRows * 8 + 1, NumberOfColumns + 2)).HorizontalAlignment = XlHAlign.xlHAlignCenter

        For i = 1 To NumberOfRows - 1
            With sheet2.Cells((i - 1) * 6 + 2, 1)
                .Value2 = sheet.Cells(i + 1, 1).Value2  'Row title
                .Interior.Color = RGB(160, 190, 230)
                .Font.Bold = True
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
            End With

            With sheet3.Cells((i - 1) * 8 + 2, 1)
                .Value2 = sheet.Cells(i + 1, 1).Value2  'Row title
                .Interior.Color = RGB(160, 190, 230)
                .Font.Bold = True
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
            End With


            sheet2.Range(sheet2.Cells((i - 1) * 6 + 2, 3), sheet2.Cells((i - 1) * 6 + 3, NumberOfColumns + 1)).Font.Bold = False
            sheet3.Range(sheet3.Cells((i - 1) * 8 + 3, 3), sheet3.Cells((i - 1) * 8 + 7, NumberOfColumns + 1)).Font.Bold = False

            sheet2.Cells((i - 1) * 6 + 2, 2).Value2 = "Observed"
            sheet2.Cells((i - 1) * 6 + 2, 2).Interior.Color = RGB(235, 200, 200)                                                                                        'Dark Pink
            sheet2.Cells((i - 1) * 6 + 2, NumberOfColumns + 2).Interior.Color = RGB(225, 190, 190)                                                                      'Lighter Dark Pink
            sheet2.Cells((i - 1) * 6 + 2, NumberOfColumns + 3).Interior.Color = RGB(235, 200, 200)                                                                      'Dark Pink
            sheet2.Cells((i - 1) * 6 + 3, 2).Value2 = "Estimated"
            sheet2.Cells((i - 1) * 6 + 3, 2).Interior.Color = RGB(242, 220, 220)                                                                                        'Light Pink
            sheet2.Cells((i - 1) * 6 + 3, NumberOfColumns + 2).Interior.Color = RGB(232, 210, 210)                                                                      'Lighter Pink
            sheet2.Cells((i - 1) * 6 + 3, NumberOfColumns + 3).Interior.Color = RGB(242, 220, 220)                                                                      'Light Pink
            sheet2.Cells((i - 1) * 6 + 4, 2).Value2 = "(E-O)^2"
            sheet2.Cells((i - 1) * 6 + 4, 2).Interior.Color = RGB(253, 233, 217)                                                                                        'Light Orange
            sheet2.Range(sheet2.Cells((i - 1) * 6 + 4, 3), sheet2.Cells((i - 1) * 6 + 4, NumberOfColumns + 1)).Interior.Color = RGB(228, 223, 236)                      'Light Lilac
            sheet2.Cells((i - 1) * 6 + 4, NumberOfColumns + 2).Interior.Color = RGB(243, 223, 207)                                                                      'Lighter Orange
            sheet2.Cells((i - 1) * 6 + 4, NumberOfColumns + 3).Interior.Color = RGB(253, 233, 217)                                                                      'Light Orange
            sheet2.Cells((i - 1) * 6 + 5, 2).Value2 = "(O-AV)^2"
            sheet2.Cells((i - 1) * 6 + 5, 2).Interior.Color = RGB(252, 213, 180)                                                                                        'Orange
            sheet2.Range(sheet2.Cells((i - 1) * 6 + 5, 3), sheet2.Cells((i - 1) * 6 + 5, NumberOfColumns + 1)).Interior.Color = RGB(204, 192, 218)                      'Lilac
            sheet2.Cells((i - 1) * 6 + 5, NumberOfColumns + 2).Interior.Color = RGB(242, 203, 170)                                                                      'Lighter Orange
            sheet2.Cells((i - 1) * 6 + 5, NumberOfColumns + 3).Interior.Color = RGB(252, 213, 180)                                                                      'Orange
            sheet2.Cells((i - 1) * 6 + 6, 2).Value2 = "ABS(O-E)"
            sheet2.Cells((i - 1) * 6 + 6, 2).Interior.Color = RGB(250, 191, 143)                                                                                        'Dark Orange
            sheet2.Range(sheet2.Cells((i - 1) * 6 + 6, 3), sheet2.Cells((i - 1) * 6 + 6, NumberOfColumns + 1)).Interior.Color = RGB(177, 160, 200)                      'Dark Lilac
            sheet2.Cells((i - 1) * 6 + 6, NumberOfColumns + 2).Interior.Color = RGB(240, 181, 133)                                                                      'Lighter Dark Orange
            sheet2.Cells((i - 1) * 6 + 6, NumberOfColumns + 3).Interior.Color = RGB(250, 191, 143)                                                                      'Dark Orange
            sheet2.Range(sheet2.Cells((i - 1) * 6 + 2, 3), sheet2.Cells((i - 1) * 6 + 6, NumberOfColumns + 8)).NumberFormat = "0.000"
            sheet2.Range(sheet2.Cells((i - 1) * 6 + 2, 3), sheet2.Cells((i - 1) * 6 + 2, NumberOfColumns + 1)).Value2 = sheet.Range(sheet.Cells(i + 1, 2), sheet.Cells(i + 1, NumberOfColumns)).Value2  'Data

            sheet3.Cells((i - 1) * 8 + 2, 2).Value2 = "Observation"
            sheet3.Cells((i - 1) * 8 + 2, 2).Interior.Color = RGB(160, 210, 95)                                                                                         'Dark Green
            sheet3.Cells((i - 1) * 8 + 2, NumberOfColumns + 2).Interior.Color = RGB(160, 210, 95)                                                                       'Dark Green
            sheet3.Range(sheet3.Cells((i - 1) * 8 + 2, 3), sheet3.Cells((i - 1) * 8 + 2, NumberOfColumns + 1)).Interior.Color = RGB(180, 220, 130)              'Light Green

            sheet3.Cells((i - 1) * 8 + 2, NumberOfColumns + 2).Value2 = "Corr."
            sheet3.Cells((i - 1) * 8 + 2, NumberOfColumns * 2 + 5).Value2 = "Com.Corr."  'Combined Correlation (The correlations between best estimations in rows)

            For j = 1 To 5
                sheet3.Cells((i - 1) * 8 + j + 2, 2).Value2 = "Estimation " + CStr(j)
                sheet3.Cells((i - 1) * 8 + j + 2, 2).Interior.Color = RGB(255, 240, 140 + (j - 1) * 20)                                                                        'Light Pink
                sheet3.Cells((i - 1) * 8 + j + 2, NumberOfColumns + 2).Interior.Color = RGB(255, 240, 140 + (j - 1) * 20)                                                      'Honey Color Tones
                sheet3.Range(sheet3.Cells((i - 1) * 8 + j + 2, 3), sheet3.Cells((i - 1) * 8 + j + 2, NumberOfColumns + 1)).Interior.Color = RGB(255, 255, 140 + (j - 1) * 20)  'Yellow Tones
            Next

            sheet3.Cells((i - 1) * 8 + 8, 2).Value2 = "Nearest Est."
            sheet3.Cells((i - 1) * 8 + 8, 2).Interior.Color = RGB(180, 220, 140)                                                                   'Dark Green
            sheet3.Cells((i - 1) * 8 + 8, NumberOfColumns + 2).Interior.Color = RGB(200, 215, 235)                                                 'Dark Green
            sheet3.Range(sheet3.Cells((i - 1) * 8 + 8, 3), sheet3.Cells((i - 1) * 8 + 8, NumberOfColumns + 1)).Interior.Color = RGB(200, 230, 170) 'Light Green

            sheet3.Range(sheet3.Cells((i - 1) * 8 + 2, 3), sheet3.Cells((i - 1) * 8 + 8, NumberOfColumns + 8)).NumberFormat = "0.000"
            sheet3.Range(sheet3.Cells((i - 1) * 8 + 2, 3), sheet3.Cells((i - 1) * 8 + 2, NumberOfColumns + 1)).Value2 = sheet.Range(sheet.Cells(i + 1, 2), sheet.Cells(i + 1, NumberOfColumns)).Value2  'Data

            sheet4.Range(sheet4.Cells((i - 1) * (NumberOfColumns - 1) + 2, 2), sheet4.Cells(i * (NumberOfColumns - 1) + 1, 2)).Value2 = excel.Transpose(sheet.Range(sheet.Cells(i + 1, 2), sheet.Cells(i + 1, NumberOfColumns)).Value2)  'Data
            sheet4.Cells((i - 1) * (NumberOfColumns - 1) + 2, 1).Value2 = ObservedData(i + 1, 1)
            sheet4.Cells((i - 1) * (NumberOfColumns - 1) + 2, 1).Interior.Color = RGB(160, 190, 230)

            For j = 0 To 4
                sheet2.Cells((i - 1) * 6 + 3, NumberOfColumns + 4 + j).Interior.Color = RGB(220 - j * 10, 230 - j * 10, 190 - j * 10)                                   'Green Tones                                                                                  
            Next
        Next

        sheet2.Columns("A:B").HorizontalAlignment = XlHAlign.xlHAlignCenter
        sheet3.Columns("A:B").HorizontalAlignment = XlHAlign.xlHAlignCenter
        sheet3.Columns("B").EntireColumn.AutoFit()

        ' Locating row and column numbers of the matrix
        For i = 1 To NumberOfColumns + 5
            sheet.Cells(NumberOfRows + 4, i + 1).Value2 = i
            sheet.Cells(NumberOfRows + 4, i + NumberOfColumns + 9).Value2 = i
        Next

        For i = 1 To NumberOfRows + 5
            sheet.Cells(i + NumberOfRows + 5, NumberOfColumns + 7).Value2 = i
            sheet.Cells(i + NumberOfRows + 5, NumberOfColumns * 2 + 15).Value2 = i
        Next

        ' Generation of the required matrices
        Dim nan(,) As Object = (sheet.Range(sheet.Cells(NumberOfRows + 9, 5), sheet.Cells(NumberOfRows * 2 + 7, NumberOfColumns + 3)).Value2)           ' nan: The observed data matrix of the station
        Dim nanWide(,) As Object = (sheet.Range(sheet.Cells(NumberOfRows + 6, 2), sheet.Cells(NumberOfRows * 2 + 10, NumberOfColumns + 6)).Value2)      ' nanWide: Expanded data matrix of the station

        Dim komsu(NumberOfRows + 5, NumberOfColumns + 5) As Integer
        Dim EstimatedData(NumberOfRows, NumberOfColumns) As Double
        Dim RemovedData(NumberOfColumns) As Double

        Dim MaxKomsu As Integer = 0
        Dim iMaxKomsu As Integer = 0
        Dim jMaxKomsu As Integer = 0
        Dim EksikVeriNo As Integer = 0
        Dim ElemanSayisi As Integer = 0   'Total number of observations in the matrix
        Dim SortedData(5000) As Double
        Dim SortedDataCoord(5000, 2) As Integer

        Dim Hold, Hold1, Hold2 As Double
        Dim HoldXY(2) As Integer

        Dim i0 As Integer               'Loop variable for the number of range clusters
        Dim i0Max As Integer = 12       'Maximum number of range clusters ***********************************************************
        Dim Dilim As Double             'Width of cluster 
        Dim DilimNo(NumberOfRows + 10, NumberOfColumns + 10) As Integer  'The order of the cluster in the data range
        Dim DilimMax As Integer         'The cluster with highest frequency
        Dim TopSay(10000) As Integer         'The frequencies of each cluster
        Dim TopDilim(10000) As Double        'The total of observations generating the frequency
        Dim SortedTopSay(10000) As Integer         'Sorted frequencies
        Dim SortedTopDilim(10000) As Double        'Sorted total observations
        Dim TopMax As Integer           'Maximum triad
        Dim UcluToplam As Integer       'The total of consequtive three clusters
        Dim UcluToplamDizisi(i0Max + 2) As Integer 'Vector for triad totals
        Dim DegerToplam As Double
        Dim DegerToplamDizisi(i0Max + 2) As Double
        Dim BestFiveEstimations(NumberOfRows, NumberOfColumns, 6) As Double   'The best 5 estimations for each observation
        Dim BestCorrelatedEstimation(NumberOfRows, NumberOfColumns) As Double 'The series producing the best correlation
        Dim BestNearestEstimation(NumberOfRows, NumberOfColumns) As Double 'The nearest estimations within the first 5 estimations
        Dim FrekansToplam As Integer
        Dim UcluMax As Integer
        Dim ksirasi As Integer
        Dim isirasi As Integer
        Dim i1, j1 As Integer
        Dim RowNumber, ColumnNumber, MinFarkRow As Integer  'Loop variable for removed rows and columns
        Dim NumberOfCells, NumberOfMissingData As Integer
        Dim MaxAkim As Double = 0
        Dim MinAkim As Double = 100000000
        Dim Estimation As Double
        Dim OnSonrasiToplam As Double
        Dim Fark1, Fark2, Fark3, Fark3Toplami, MinFark, ObservedRowAverage, EstimatedRowAverage, SixthRowAverage As Double
        Dim SatirToplami1, SatirToplami2, RMSE, RowMax, RowMin As Double
        Dim Ortalama1, Ortalama2, MissingDataRate As Double
        Dim CorrelationRange1, CorrelationRange2 As Range

        'Conditional Coloring of Observed and Estimated Columns in Cumulative Statistics and Graphs file
        rng = sheet4.Range(sheet4.Cells(2, 2), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 3))
        rng.FormatConditions.Delete()
        cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)

        sheet.Range(sheet.Cells(NumberOfRows + 6, 2), sheet.Cells(NumberOfRows * 2 + 10, NumberOfColumns + 6)).NumberFormat = "0.00"

        For RowNumber = 1 To NumberOfRows - 1
            'Removing observed values in rows:
            For ColumnNumber = 1 To NumberOfColumns - 1
                RemovedData(ColumnNumber) = nan(RowNumber, ColumnNumber)
                nan(RowNumber, ColumnNumber) = Nothing
                nanWide(RowNumber + 3, ColumnNumber + 3) = Nothing
                If ColumnNumber < 4 Then nanWide(RowNumber + 2, ColumnNumber + NumberOfColumns + 2) = Nothing
                If ColumnNumber > 12 Then nanWide(RowNumber + 4, ColumnNumber - NumberOfColumns + 1) = Nothing
            Next

100:        ' Determination of number of existing observed neighbors for each cell
            For i = 4 To NumberOfRows + 2
                For j = 4 To NumberOfColumns + 2
                    komsu(i, j) = 0
                    For k = -3 To 3
                        For m = -3 To 3
                            If nanWide(i + k, j + m) IsNot Nothing Then komsu(i, j) = komsu(i, j) + 1
                            If nanWide(i + k, j + m) = 0 And nanWide(i + k, j + m) IsNot Nothing Then komsu(i, j) = komsu(i, j) + 1
                        Next
                    Next
                Next
            Next

            'Determination of the first cell to be estimated (The cell with maximum number of observed neighbors):

            'Copying observations to the vector to be sorted
            ElemanSayisi = 0

            For i = 4 To NumberOfRows + 2
                For j = 4 To NumberOfColumns + 2
                    ElemanSayisi = ElemanSayisi + 1
                    SortedData(ElemanSayisi) = 0
                    SortedDataCoord(ElemanSayisi, 1) = 0
                    SortedDataCoord(ElemanSayisi, 2) = 0
                Next
            Next

            ElemanSayisi = 0
            For i = 4 To NumberOfRows + 2
                For j = 4 To NumberOfColumns + 2
                    If nanWide(i, j) IsNot Nothing Then
                        ElemanSayisi = ElemanSayisi + 1
                        SortedData(ElemanSayisi) = nanWide(i, j)
                        SortedDataCoord(ElemanSayisi, 1) = i
                        SortedDataCoord(ElemanSayisi, 2) = j
                    End If
                Next
            Next

            'Bubble sorting observations:
            For i = 0 To ElemanSayisi - 1
                For j = i + 1 To ElemanSayisi
                    If SortedData(j) < SortedData(i) Then
                        Hold = SortedData(i)
                        HoldXY(1) = SortedDataCoord(i, 1)
                        HoldXY(2) = SortedDataCoord(i, 2)
                        SortedData(i) = SortedData(j)
                        SortedDataCoord(i, 1) = SortedDataCoord(j, 1)
                        SortedDataCoord(i, 2) = SortedDataCoord(j, 2)
                        SortedData(j) = Hold
                        SortedDataCoord(j, 1) = HoldXY(1)
                        SortedDataCoord(j, 2) = HoldXY(2)
                    End If
                Next
            Next

            MaxKomsu = 0
            EksikVeriNo = EksikVeriNo + 1
            OnSonrasiToplam = 0

            For j = 4 To NumberOfColumns + 2
                If nanWide(RowNumber + 3, j) = 0.0 And nanWide(RowNumber + 3, j) IsNot Nothing Then GoTo 105
                If nanWide(RowNumber + 3, j) = Nothing And komsu(RowNumber + 3, j) > MaxKomsu Then
                    MaxKomsu = komsu(RowNumber + 3, j)
                    iMaxKomsu = RowNumber + 3
                    jMaxKomsu = j
                End If
105:        Next

            MaxAkim = 0.0
            MinAkim = 100000000

            'Determination of Maximum and Minimum observations
            For i = 4 To NumberOfRows + 2
                For j = 4 To NumberOfColumns + 2
                    If nanWide(i, j) = 0 And MinAkim > 0 Then
                        MinAkim = 0
                        Continue For
                    End If
                    If nanWide(i, j) = 0 Then GoTo 110
                    If nanWide(i, j) = Nothing Then Continue For
110:                If nanWide(i, j) > MaxAkim Then MaxAkim = nanWide(i, j)
                    If nanWide(i, j) < MinAkim Then MinAkim = nanWide(i, j)
                Next
            Next

            'Locating the clusters on the table below
            sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 1).Value2 = "i"
            sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 2).Value2 = "j"
            sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3) + 1, 1).Value2 = ObservedData(iMaxKomsu - 2, 1)
            sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3) + 1, 2).Value2 = ObservedData(1, jMaxKomsu - 2)
            sheet.Range(sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 1), sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3) + 1, 2)).HorizontalAlignment = XlHAlign.xlHAlignCenter

            For i = 2 To i0Max
                sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i + 3).Value2 = i
            Next

            For i = 1 To i0Max
                sheet.Cells(i + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 4).Value2 = i
            Next

            For i = 2 To i0Max
                sheet.Cells(i + NumberOfRows * 2 + 12 + (EksikVeriNo - 1) * (i0Max + 3), 8 + i0Max).Value2 = i
            Next

            rng = sheet.Range(sheet.Cells(NumberOfRows * 2 + 14 + (EksikVeriNo - 1) * (i0Max + 3), 5 + i0Max), sheet.Cells(NumberOfRows * 2 + 13 + i0Max + (EksikVeriNo - 1) * (i0Max + 3), 6 + i0Max))
            cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)

            sheet.Range(sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 1), sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 2)).Interior.Color = RGB(190, 215, 245)
            sheet.Range(sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3) + 1, 1), sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3) + 1, 2)).Interior.Color = RGB(255, 255, 204)
            sheet.Range(sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 5), sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 3)).Interior.Color = RGB(190, 215, 245)
            sheet.Range(sheet.Cells(NumberOfRows * 2 + 14 + (EksikVeriNo - 1) * (i0Max + 3), 4), sheet.Cells(i0Max + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 4)).Interior.Color = RGB(190, 215, 245)
            sheet.Range(sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 5), sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 6)).Interior.Color = RGB(190, 215, 245)
            sheet.Range(sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 8), sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 11)).Interior.Color = RGB(190, 215, 245)
            sheet.Range(sheet.Cells(NumberOfRows * 2 + 14 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 8), sheet.Cells(i0Max + NumberOfRows * 2 + 12 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 8)).Interior.Color = RGB(190, 215, 245)

            For i0 = 2 To i0Max
                'Dividing the data into clusters and determination of the cluster numbers and writing on the table on the right:
                'The number of observations are equal in each cluster:
                Dilim = ElemanSayisi / i0
                For i = 1 To ElemanSayisi
                    DilimNo(SortedDataCoord(i, 1), SortedDataCoord(i, 2)) = Int(i / Dilim) + 1
                    If (Int(i / Dilim) + 1) > i0 Then DilimNo(SortedDataCoord(i, 1), SortedDataCoord(i, 2)) = i0
                Next

                'Dividing the data into clusters and determination of the cluster numbers and writing on the table on the right:
                'Cluster Ranges are equal:
                'Dilim = (MaxAkim - MinAkim) / i0
                'For i = 1 To NumberOfRows + 5
                'For j = 1 To NumberOfColumns + 5
                'If nanWide(i, j) = 0 Then goto 115
                'If nanWide(i, j) = Nothing Then
                'DilimNo(i, j) = 0
                'sheet.Cells(i + NumberOfRows + 5, j + NumberOfColumns + 9).Value2 = DilimNo(i, j)    
                'Continue For
                'End If
115:            'DilimNo(i, j) = Int((nanWide(i, j) - MinAkim) / Dilim) + 1
                'If nanWide(i, j) = MaxAkim Then DilimNo(i, j) = DilimNo(i, j) - 1
                'If DilimNo(i, j) > i0 Then DilimNo(i, j) = i0
                'sheet.Cells(i + NumberOfRows + 5, j + NumberOfColumns + 9).Value2 = DilimNo(i, j)    
                'Next
                'Next


                For i = 1 To NumberOfRows + 5
                    For j = 1 To NumberOfColumns + 5
                        If nanWide(i, j) = Nothing Then DilimNo(i, j) = 0
                        If nanWide(i, j) = MaxAkim Then DilimNo(i, j) = i0
                    Next
                Next

                'Determination of cluster numbers of the right and left portions of the extended matrix
                For i = 5 To NumberOfRows + 3
                    For j = 1 To 3
                        DilimNo(i, j) = DilimNo(i - 1, j + NumberOfColumns - 1)
                    Next
                Next

                For i = 3 To NumberOfRows + 1
                    For j = NumberOfColumns + 3 To NumberOfColumns + 5
                        DilimNo(i, j) = DilimNo(i + 1, j - (NumberOfColumns - 1))
                    Next
                Next


                'Searching the cluster pairs in the dataset and determination of the cluster of the missing data
                For i = 0 To i0Max
                    TopSay(i) = 0
                    TopDilim(i) = 0.0
                Next


                For i1 = 4 To NumberOfRows + 2
                    For j1 = 4 To NumberOfColumns + 2
                        For i = -3 To 2
                            For j = -3 To 2
                                'Diagonal (Left top - right bottom)
                                If (DilimNo(i1 + i, j1 + j) = DilimNo(iMaxKomsu + i, jMaxKomsu + j)) And (DilimNo(i1 + i, j1 + j) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu + j) > 0) And DilimNo(i1, j1) > 0 Then
                                    If (DilimNo(i1 + i + 1, j1 + j + 1) = DilimNo(iMaxKomsu + i + 1, jMaxKomsu + j + 1)) And (DilimNo(i1 + i + 1, j1 + j + 1) > 0 And DilimNo(iMaxKomsu + i + 1, jMaxKomsu + j + 1) > 0) Then
                                        TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                        TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                                    End If
                                End If


                                'Diagonal (Right top - left bottom)
                                If (DilimNo(i1 + i, j1 + j + 1) = DilimNo(iMaxKomsu + i, jMaxKomsu + j + 1)) And (DilimNo(i1 + i, j1 + j + 1) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu + j + 1) > 0) And DilimNo(i1, j1) > 0 Then
                                    If (DilimNo(i1 + i + 1, j1 + j) = DilimNo(iMaxKomsu + i + 1, jMaxKomsu + j)) And (DilimNo(i1 + i + 1, j1 + j) > 0 And DilimNo(iMaxKomsu + i + 1, jMaxKomsu + j) > 0) Then
                                        TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                        TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                                    End If
                                End If

                                'Horizontal
                                If (DilimNo(i1 + i, j1 + j) = DilimNo(iMaxKomsu + i, jMaxKomsu + j)) And (DilimNo(i1 + i, j1 + j) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu + j) > 0) And DilimNo(i1, j1) > 0 Then
                                    If (DilimNo(i1 + i, j1 + j + 1) = DilimNo(iMaxKomsu + i, jMaxKomsu + j + 1)) And (DilimNo(i1 + i, j1 + j + 1) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu + j + 1) > 0) Then
                                        TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                        TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                                    End If
                                End If

                                'Square
                                If (DilimNo(i1 + i, j1 + j) = DilimNo(iMaxKomsu + i, jMaxKomsu + j)) And (DilimNo(i1 + i, j1 + j) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu + j) > 0) And DilimNo(i1, j1) > 0 Then
                                    If (DilimNo(i1 + i, j1 + j + 1) = DilimNo(iMaxKomsu + i, jMaxKomsu + j + 1)) And (DilimNo(i1 + i, j1 + j + 1) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu + j + 1) > 0) Then
                                        If (DilimNo(i1 + i + 1, j1 + j + 1) = DilimNo(iMaxKomsu + i + 1, jMaxKomsu + j + 1)) And (DilimNo(i1 + i + 1, j1 + j + 1) > 0 And DilimNo(iMaxKomsu + i + 1, jMaxKomsu + j + 1) > 0) Then
                                            If (DilimNo(i1 + i + 1, j1 + j) = DilimNo(iMaxKomsu + i + 1, jMaxKomsu + j)) And (DilimNo(i1 + i + 1, j1 + j) > 0 And DilimNo(iMaxKomsu + i + 1, jMaxKomsu + j) > 0) Then
                                                TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                                TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                                            End If
                                        End If
                                    End If
                                End If
                            Next

                            'The horizontal line at the bottom
                            For j = -3 To 2
                                If (DilimNo(i1 + 3, j1 + j) = DilimNo(iMaxKomsu + 3, jMaxKomsu + j)) And (DilimNo(i1 + 3, j1 + j) > 0 And DilimNo(iMaxKomsu + 3, jMaxKomsu + j) > 0) And DilimNo(i1, j1) > 0 Then
                                    If (DilimNo(i1 + 3, j1 + j + 1) = DilimNo(iMaxKomsu + 3, jMaxKomsu + j + 1)) And (DilimNo(i1 + 3, j1 + j + 1) > 0 And DilimNo(iMaxKomsu + 3, jMaxKomsu + j + 1) > 0) Then
                                        TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                        TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                                    End If
                                End If
                            Next

                        Next

                        'Three nearest diagonal pairs (missing data in the middle)
                        For j = -1 To 1
                            If (DilimNo(i1 - 1, j1 + j) = DilimNo(iMaxKomsu - 1, jMaxKomsu + j)) And (DilimNo(i1 - 1, j1 + j) > 0 And DilimNo(iMaxKomsu - 1, jMaxKomsu + j) > 0) And DilimNo(i1, j1) > 0 Then
                                If (DilimNo(i1 + 1, j1 - j) = DilimNo(iMaxKomsu + 1, jMaxKomsu - j)) And (DilimNo(i1 + 1, j1 - j) > 0 And DilimNo(iMaxKomsu + 1, jMaxKomsu - j) > 0) Then
                                    TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                    TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                                End If
                            End If
                        Next

                        'Three nearest horizontal pairs (missing data in the middle)
                        For i = -1 To 1
                            If (DilimNo(i1 + i, j1 - 1) = DilimNo(iMaxKomsu + i, jMaxKomsu - 1)) And (DilimNo(i1 + i, j1 - 1) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu - 1) > 0) And DilimNo(i1, j1) > 0 Then
                                If (DilimNo(i1 + i, j1 + 1) = DilimNo(iMaxKomsu + i, jMaxKomsu + 1)) And (DilimNo(i1 + i, j1 + 1) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu + 1) > 0) Then
                                    TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                    TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                                End If
                            End If
                        Next

                        '6 inclined diagonal pairs
                        If (DilimNo(i1 - 1, j1 - 1) = DilimNo(iMaxKomsu - 1, jMaxKomsu - 1)) And (DilimNo(i1 - 1, j1 - 1) > 0 And DilimNo(iMaxKomsu - 1, jMaxKomsu - 1) > 0) And DilimNo(i1, j1) > 0 Then
                            If (DilimNo(i1 + 1, j1) = DilimNo(iMaxKomsu + 1, jMaxKomsu)) And (DilimNo(i1 + 1, j1) > 0 And DilimNo(iMaxKomsu + 1, jMaxKomsu) > 0) Then
                                TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                            End If
                            If (DilimNo(i1, j1 + 1) = DilimNo(iMaxKomsu, jMaxKomsu + 1)) And (DilimNo(i1, j1 + 1) > 0 And DilimNo(iMaxKomsu, jMaxKomsu + 1) > 0) Then
                                TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                            End If
                        End If

                        If (DilimNo(i1 - 1, j1) = DilimNo(iMaxKomsu - 1, jMaxKomsu)) And (DilimNo(i1 - 1, j1) > 0 And DilimNo(iMaxKomsu - 1, jMaxKomsu) > 0) And DilimNo(i1, j1) > 0 Then
                            If (DilimNo(i1 + 1, j1 - 1) = DilimNo(iMaxKomsu + 1, jMaxKomsu - 1)) And (DilimNo(i1 + 1, j1 - 1) > 0 And DilimNo(iMaxKomsu + 1, jMaxKomsu - 1) > 0) Then
                                TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                            End If
                            If (DilimNo(i1 + 1, j1 + 1) = DilimNo(iMaxKomsu + 1, jMaxKomsu + 1)) And (DilimNo(i1 + 1, j1 + 1) > 0 And DilimNo(iMaxKomsu + 1, jMaxKomsu + 1) > 0) Then
                                TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                            End If
                        End If

                        If (DilimNo(i1 - 1, j1 + 1) = DilimNo(iMaxKomsu - 1, jMaxKomsu + 1)) And (DilimNo(i1 - 1, j1 + 1) > 0 And DilimNo(iMaxKomsu - 1, jMaxKomsu + 1) > 0) And DilimNo(i1, j1) > 0 Then
                            If (DilimNo(i1 + 1, j1) = DilimNo(iMaxKomsu + 1, jMaxKomsu)) And (DilimNo(i1 + 1, j1) > 0 And DilimNo(iMaxKomsu + 1, jMaxKomsu) > 0) Then
                                TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                            End If
                            If (DilimNo(i1, j1 - 1) = DilimNo(iMaxKomsu, jMaxKomsu - 1)) And (DilimNo(i1, j1 - 1) > 0 And DilimNo(iMaxKomsu, jMaxKomsu - 1) > 0) Then
                                TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                            End If
                        End If

                        'All vertiacla pairs in the same column of the missing data
                        For i = -3 To 2
                            If i = 0 Then Continue For
                            If (DilimNo(i1 + i, j1) = DilimNo(iMaxKomsu + i, jMaxKomsu)) And (DilimNo(i1 + i, j1) > 0 And DilimNo(iMaxKomsu + i, jMaxKomsu) > 0) And DilimNo(i1, j1) > 0 Then
                                For k = i + 1 To 3
                                    If k = 0 Then Continue For
                                    If i = -1 And k = 1 Then Continue For
                                    If (DilimNo(i1 + k, j1) = DilimNo(iMaxKomsu + k, jMaxKomsu)) And (DilimNo(i1 + k, j1) > 0 And DilimNo(iMaxKomsu + k, jMaxKomsu) > 0) Then
                                        TopSay(DilimNo(i1, j1)) = TopSay(DilimNo(i1, j1)) + 1
                                        TopDilim(DilimNo(i1, j1)) = TopDilim(DilimNo(i1, j1)) + nanWide(i1, j1)
                                    End If
                                Next
                            End If
                        Next

                    Next
                Next

                'Writing the frequencies to the table below
                For i = 1 To i0
                    sheet.Cells(i + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 3 + i0).Value2 = TopSay(i)
                Next

                sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 5 + i0Max).Value2 = "Min"
                sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 6 + i0Max).Value2 = "Max"
                sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 8 + i0Max).Value2 = "Group #"
                sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 9 + i0Max).Value2 = "Estimation"
                sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 10 + i0Max).Value2 = "Min"
                sheet.Cells(NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), 11 + i0Max).Value2 = "Max"

                'Conditional coloring according to frequencies
                rng = sheet.Range(sheet.Cells(NumberOfRows * 2 + 14 + (EksikVeriNo - 1) * (i0Max + 3), 3 + i0), sheet.Cells(NumberOfRows * 2 + 13 + i0 + (EksikVeriNo - 1) * (i0Max + 3), 3 + i0))
                rng.FormatConditions.Delete()
                cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)

                'Determination of three consequtive clusters with the highest total frequency
                If i0 > 2 Then
                    DilimMax = 0
                    TopMax = 0
                    For i = 1 To i0Max - 2
                        UcluToplam = 0
                        For j = 0 To 2
                            UcluToplam = UcluToplam + TopSay(i + j)
                        Next
                        If UcluToplam > TopMax Then
                            TopMax = UcluToplam
                            UcluMax = 0
                            DegerToplam = 0
                            FrekansToplam = 0
                            isirasi = i
                            For k = i To i + 2
                                If TopSay(k) > UcluMax Then
                                    UcluMax = TopSay(k)
                                    ksirasi = k
                                End If
                                DegerToplam = DegerToplam + TopDilim(k)
                                FrekansToplam = FrekansToplam + TopSay(k)
                            Next k
                            DilimMax = ksirasi
                            If TopSay(i + 1) = TopSay(i) And UcluMax = TopSay(i + 1) Then DilimMax = i + 1
                            If TopSay(i + 1) = TopSay(i + 2) And UcluMax = TopSay(i + 1) Then DilimMax = i + 1
                        End If
                    Next
                End If

                'Writing the estimated data on the table to the right
                If i0 > 2 Then
                    If DilimMax = 0 Then DilimMax = 1
                    Estimation = DegerToplam / FrekansToplam
                    If i0 > 9 Then OnSonrasiToplam = OnSonrasiToplam + Estimation
                    sheet.Cells(NumberOfRows * 2 + 12 + i0 + (EksikVeriNo - 1) * (i0Max + 3), 9 + i0Max).Value2 = Estimation

                    'Writing cluster triad range to the table on the right
                    sheet.Range(sheet.Cells(NumberOfRows * 2 + 12 + i0 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 9), sheet.Cells(NumberOfRows * 2 + 12 + i0 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 11)).NumberFormat = "0.0"
                    sheet.Cells(NumberOfRows * 2 + 12 + i0 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 10).Value2 = SortedData(Int(Dilim * (isirasi - 1)) + 1)
                    sheet.Cells(NumberOfRows * 2 + 12 + i0 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 11).Value2 = SortedData(Int(Dilim * (isirasi + 2)))
                    'sheet.Cells(NumberOfRows * 2 + 12 + i0 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 10).Value2 = MinAkim + Dilim * (isirasi - 1)  'For equal cluster ranges
                    'sheet.Cells(NumberOfRows * 2 + 12 + i0 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 11).Value2 = MinAkim + Dilim * (isirasi + 2)  'For equal cluster ranges
                End If


            Next   'i0 loop


            'Writing cluster ranges to the right
            sheet.Range(sheet.Cells(1 + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 5), sheet.Cells(i0Max + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 6)).NumberFormat = "0.0"
            For i = 1 To i0Max
                sheet.Cells(i + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 5).Value2 = SortedData(Dilim * (i - 1) + 1)
                sheet.Cells(i + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 6).Value2 = SortedData(Dilim * i)
                'sheet.Cells(i + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 5).Value2 = MinAkim + Dilim * (i - 1)   'For equal cluster ranges
                'sheet.Cells(i + NumberOfRows * 2 + 13 + (EksikVeriNo - 1) * (i0Max + 3), i0Max + 6).Value2 = MinAkim + Dilim * i         'For equal cluster ranges
            Next

            'Determination of the first five cluster triads with the highest total frequencies for i0Max value (the last column of the frequency table)
            'Determination of three consequtive clusters with the highest total frequency
            'DilimMax = 0
            'TopMax = 0
            For i = 1 To i0Max - 2
                UcluToplamDizisi(i) = 0
                DegerToplamDizisi(i) = 0
                For j = 0 To 2
                    UcluToplamDizisi(i) = UcluToplamDizisi(i) + TopSay(i + j)
                    DegerToplamDizisi(i) = DegerToplamDizisi(i) + TopDilim(i + j)
                Next
            Next

            'Sorting the total frequency triads and observed totals for the cluster triads:
            For i = 1 To i0Max - 2
                For j = i + 1 To i0Max - 1
                    If UcluToplamDizisi(j) > UcluToplamDizisi(i) Then
                        Hold1 = UcluToplamDizisi(i)
                        Hold2 = DegerToplamDizisi(i)
                        UcluToplamDizisi(i) = UcluToplamDizisi(j)
                        DegerToplamDizisi(i) = DegerToplamDizisi(j)
                        UcluToplamDizisi(j) = Hold1
                        DegerToplamDizisi(j) = Hold2
                    End If
                Next
            Next

            'Sorting the total frequencies and observed totals of each cluster:
            For i = 1 To i0Max
                SortedTopSay(i) = TopSay(i)
                SortedTopDilim(i) = TopDilim(i)
            Next

            For i = 1 To i0Max - 1
                For j = i + 1 To i0Max
                    If SortedTopSay(j) > SortedTopSay(i) Then
                        Hold1 = SortedTopSay(i)
                        Hold2 = SortedTopDilim(i)
                        SortedTopSay(i) = SortedTopSay(j)
                        SortedTopDilim(i) = SortedTopDilim(j)
                        SortedTopSay(j) = Hold1
                        SortedTopDilim(j) = Hold2
                    End If
                Next
            Next

            'Generating the best first 5 estimation series
            For i = 1 To 5
                'BestFiveEstimations(RowNumber, jMaxKomsu - 3, i) = DegerToplamDizisi(i) / UcluToplamDizisi(i) 'For three consequtive clusters
                BestFiveEstimations(RowNumber, jMaxKomsu - 3, i) = SortedTopDilim(i) / SortedTopSay(i)         'For a single cluster
                If DegerToplamDizisi(i) = 0 Then BestFiveEstimations(RowNumber, jMaxKomsu - 3, i) = 0 'For three consequtive clusters
                If SortedTopDilim(i) = 0 Then BestFiveEstimations(RowNumber, jMaxKomsu - 3, i) = 0 'For a single cluster
            Next

            'Imputation of missing values od locating on the table
            If DilimMax = 0 Then DilimMax = 1
            nanWide(iMaxKomsu, jMaxKomsu) = OnSonrasiToplam / (i0Max - 9)

            If sheet.Cells(NumberOfRows + 5 + iMaxKomsu, 1 + jMaxKomsu).Value2 = Nothing Then sheet.Cells(NumberOfRows + 5 + iMaxKomsu, 1 + jMaxKomsu).Value2 = nanWide(iMaxKomsu, jMaxKomsu)
            If jMaxKomsu = 4 Or jMaxKomsu = 5 Or jMaxKomsu = 6 Then
                nanWide(iMaxKomsu - 1, jMaxKomsu + NumberOfColumns - 1) = DilimMax
                If sheet.Cells(NumberOfRows + 4 + iMaxKomsu, NumberOfColumns + jMaxKomsu).Value2 = Nothing Then sheet.Cells(NumberOfRows + 4 + iMaxKomsu, NumberOfColumns + jMaxKomsu).Value2 = nanWide(iMaxKomsu, jMaxKomsu)
            End If

            If jMaxKomsu = NumberOfColumns Or jMaxKomsu = NumberOfColumns + 1 Or jMaxKomsu = NumberOfColumns + 2 Then
                nanWide(iMaxKomsu + 1, jMaxKomsu - (NumberOfColumns - 1)) = DilimMax
                If sheet.Cells(NumberOfRows + 6 + iMaxKomsu, jMaxKomsu - (NumberOfColumns - 2)).Value2 = Nothing Then sheet.Cells(NumberOfRows + 6 + iMaxKomsu, jMaxKomsu - (NumberOfColumns - 2)).Value2 = nanWide(iMaxKomsu, jMaxKomsu)
            End If

            EstimatedData(iMaxKomsu - 3, jMaxKomsu - 3) = nanWide(iMaxKomsu, jMaxKomsu)
            'Writing the estimation to the Comparisons file
            sheet2.Cells((RowNumber - 1) * 6 + 3, jMaxKomsu - 1).Value2 = nanWide(iMaxKomsu, jMaxKomsu)

            For j = 4 To NumberOfColumns + 2
                If DilimNo(RowNumber + 3, j) = 0 Then GoTo 100
            Next

            'Determination of the estimation series giving the best correlation:
            For j = 1 To 5
                MinFark = 10000
                For i = 1 To NumberOfColumns - 1
                    sheet3.Cells((RowNumber - 1) * 8 + 2 + j, i + 2) = BestFiveEstimations(RowNumber, i, j)
                Next
                CorrelationRange1 = sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2, 3), sheet3.Cells((RowNumber - 1) * 8 + 2, NumberOfColumns + 1))
                CorrelationRange2 = sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2 + j, 3), sheet3.Cells((RowNumber - 1) * 8 + 2 + j, NumberOfColumns + 1))
                SatirToplami1 = excel.WorksheetFunction.Sum(sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2, 3), sheet3.Cells((RowNumber - 1) * 8 + 2, NumberOfColumns + 1)))
                SatirToplami2 = excel.WorksheetFunction.Sum(sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2 + j, 3), sheet3.Cells((RowNumber - 1) * 8 + 2 + j, NumberOfColumns + 1)))
                If SatirToplami1 = Nothing Then Continue For
                If SatirToplami2 = 0 Then Continue For 'If all estimations are zero, correlation is undefined
                sheet3.Cells((RowNumber - 1) * 8 + 2 + j, NumberOfColumns + 2).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)
            Next

            For j = 1 To NumberOfColumns - 1
                MinFark = 10000
                If ObservedData(RowNumber + 1, j + 1) = 0 And ObservedData(RowNumber + 1, j + 1) IsNot Nothing Then GoTo 140
                If ObservedData(RowNumber + 1, j + 1) = Nothing Then GoTo 145
140:            For i = 1 To 5
                    Fark1 = Math.Abs(BestFiveEstimations(RowNumber, j, i) - ObservedData(RowNumber + 1, j + 1))
                    If Fark1 < MinFark Then
                        MinFark = Fark1
                        MinFarkRow = i
                        BestCorrelatedEstimation(RowNumber, j) = BestFiveEstimations(RowNumber, j, i)
                    End If
                Next
                sheet3.Cells((RowNumber - 1) * 8 + 8, j + 2).Value2 = BestCorrelatedEstimation(RowNumber, j)
                sheet3.Cells((RowNumber - 1) * 8 + MinFarkRow + 2, j + 2).Interior.Color = RGB(215, 235, 190)
                sheet4.Cells((RowNumber - 1) * (NumberOfColumns - 1) + j + 1, 3) = BestCorrelatedEstimation(RowNumber, j)
145:        Next

            For j = 1 To NumberOfColumns - 1

                If ObservedData(RowNumber + 1, j + 1) = 0 And ObservedData(RowNumber + 1, j + 1) IsNot Nothing Then GoTo 146
                If ObservedData(RowNumber + 1, j + 1) = Nothing Then GoTo 147
146:            For k = 2 To 5
                    MinFark = 10000
                    For i = 1 To k
                        Fark1 = Math.Abs(BestFiveEstimations(RowNumber, j, i) - ObservedData(RowNumber + 1, j + 1))
                        If Fark1 < MinFark Then
                            MinFark = Fark1
                            BestNearestEstimation(RowNumber, j) = BestFiveEstimations(RowNumber, j, i)
                        End If
                    Next
                    sheet3.Cells((RowNumber - 1) * 8 + k + 2, j + NumberOfColumns + 5).Value2 = BestNearestEstimation(RowNumber, j)
                    sheet4.Cells((RowNumber - 1) * (NumberOfColumns - 1) + j + 1, 5 + k) = BestNearestEstimation(RowNumber, j)
                Next
147:        Next

            SatirToplami1 = excel.WorksheetFunction.Sum(sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2, 3), sheet3.Cells((RowNumber - 1) * 8 + 2, NumberOfColumns + 1))) 'Observed line total
            If SatirToplami1 = Nothing Then GoTo 148
            For j = 2 To 5
                CorrelationRange1 = sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2, 3), sheet3.Cells((RowNumber - 1) * 8 + 2, NumberOfColumns + 1))
                CorrelationRange2 = sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2 + j, NumberOfColumns + 6), sheet3.Cells((RowNumber - 1) * 8 + 2 + j, NumberOfColumns * 2 + 4))
                SatirToplami2 = excel.WorksheetFunction.Sum(sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2 + j, NumberOfColumns + 6), sheet3.Cells((RowNumber - 1) * 8 + 2 + j, NumberOfColumns * 2 + 4))) 'Estimated line total
                If SatirToplami2 = 0 Then Continue For 'If all estimations are zero, correlation is undefined
                sheet3.Cells((RowNumber - 1) * 8 + 2 + j, NumberOfColumns * 2 + 5).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)
            Next

148:        CorrelationRange1 = sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 2, 3), sheet3.Cells((RowNumber - 1) * 8 + 2, NumberOfColumns + 1))              'Observed Data
            CorrelationRange2 = sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 8, 3), sheet3.Cells((RowNumber - 1) * 8 + 8, NumberOfColumns + 1))              'Nearest Estimation Data
            If SatirToplami1 = Nothing Then GoTo 150
            SatirToplami2 = excel.WorksheetFunction.Sum(sheet3.Range(sheet3.Cells((RowNumber - 1) * 8 + 8, 3), sheet3.Cells((RowNumber - 1) * 8 + 8, NumberOfColumns + 1)))
            If SatirToplami2 = 0 Then GoTo 150 'If all estimations are zero, correlation is undefined
            sheet3.Cells((RowNumber - 1) * 8 + 8, NumberOfColumns + 2).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)

150:        'Conditional coloring of Observed and Estimated rows in Comparisons file
            rng = sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 2, 3), sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 1))
            rng.FormatConditions.Delete()
            cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)

            'Writing the calculations to Comparisons file:
            sheet2.Cells(1, NumberOfColumns + 2).Value2 = "Average"
            SatirToplami1 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 2, 3), sheet2.Cells((RowNumber - 1) * 6 + 2, NumberOfColumns + 1)))
            If SatirToplami1 = Nothing Then GoTo 160
            ObservedRowAverage = excel.WorksheetFunction.Average(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 2, 3), sheet2.Cells((RowNumber - 1) * 6 + 2, NumberOfColumns + 1)))
160:        EstimatedRowAverage = excel.WorksheetFunction.Average(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 3, 3), sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 1)))
            If ObservedRowAverage = Nothing Then ObservedRowAverage = 0
            sheet2.Cells((RowNumber - 1) * 6 + 2, NumberOfColumns + 2).Value2 = ObservedRowAverage
            sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 2).Value2 = EstimatedRowAverage

            For j = 3 To NumberOfColumns + 1
                If ObservedData(RowNumber + 1, j - 1) = 0 And ObservedData(RowNumber + 1, j - 1) IsNot Nothing Then GoTo 165
                If ObservedData(RowNumber + 1, j - 1) = Nothing Then Continue For
165:            Fark1 = ObservedData(RowNumber + 1, j - 1) - EstimatedData(RowNumber, j - 2)
                sheet2.Cells((RowNumber - 1) * 6 + 4, j).Value2 = Math.Pow(Fark1, 2)
                Fark2 = ObservedData(RowNumber + 1, j - 1) - ObservedRowAverage
                sheet2.Cells((RowNumber - 1) * 6 + 5, j).Value2 = Math.Pow(Fark2, 2)
                sheet2.Cells((RowNumber - 1) * 6 + 6, j).Value2 = Math.Abs(Fark1)
            Next

            SatirToplami1 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 4, 3), sheet2.Cells((RowNumber - 1) * 6 + 4, NumberOfColumns + 1)))
            SatirToplami2 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 5, 3), sheet2.Cells((RowNumber - 1) * 6 + 5, NumberOfColumns + 1)))
            sheet2.Cells((RowNumber - 1) * 6 + 4, NumberOfColumns + 2).Value2 = SatirToplami1 / (NumberOfColumns - 1)
            sheet2.Cells((RowNumber - 1) * 6 + 5, NumberOfColumns + 2).Value2 = SatirToplami2 / (NumberOfColumns - 1)

            sheet2.Cells(1, NumberOfColumns + 3).Value2 = "Sum"

            SatirToplami1 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 2, 3), sheet2.Cells((RowNumber - 1) * 6 + 2, NumberOfColumns + 1)))
            SatirToplami2 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 3, 3), sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 1)))
            If SatirToplami1 = Nothing Then SatirToplami1 = 0
            sheet2.Cells((RowNumber - 1) * 6 + 2, NumberOfColumns + 3).Value2 = SatirToplami1
            sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 3).Value2 = SatirToplami2

            SatirToplami1 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 4, 3), sheet2.Cells((RowNumber - 1) * 6 + 4, NumberOfColumns + 1)))
            SatirToplami2 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 5, 3), sheet2.Cells((RowNumber - 1) * 6 + 5, NumberOfColumns + 1)))
            If SatirToplami1 = Nothing Then SatirToplami1 = 0
            sheet2.Cells((RowNumber - 1) * 6 + 4, NumberOfColumns + 3).Value2 = SatirToplami1
            sheet2.Cells((RowNumber - 1) * 6 + 5, NumberOfColumns + 3).Value2 = SatirToplami2

            SatirToplami1 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 6, 3), sheet2.Cells((RowNumber - 1) * 6 + 6, NumberOfColumns + 1)))
            sheet2.Cells((RowNumber - 1) * 6 + 6, NumberOfColumns + 3).Value2 = SatirToplami1

            sheet2.Cells(1, NumberOfColumns + 2).Value2 = "Average"
            SatirToplami1 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 6, 3), sheet2.Cells((RowNumber - 1) * 6 + 6, NumberOfColumns + 1)))
            If SatirToplami1 = Nothing Then GoTo 170
            SixthRowAverage = excel.WorksheetFunction.Average(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 6, 3), sheet2.Cells((RowNumber - 1) * 6 + 6, NumberOfColumns + 1)))
            sheet2.Cells((RowNumber - 1) * 6 + 6, NumberOfColumns + 2).Value2 = SixthRowAverage

170:        sheet2.Cells(1, NumberOfColumns + 4).Value2 = "Corr."
            CorrelationRange1 = sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 2, 3), sheet2.Cells((RowNumber - 1) * 6 + 2, NumberOfColumns + 1))
            CorrelationRange2 = sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 3, 3), sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 1))
            SatirToplami1 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 2, 3), sheet2.Cells((RowNumber - 1) * 6 + 2, NumberOfColumns + 1)))
            SatirToplami2 = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 3, 3), sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 1)))
            If SatirToplami1 = Nothing Then GoTo 200 'If all observations in the row are missing correlation is undefined
            If SatirToplami2 = 0 Then GoTo 200 'If all estimations are zero, correlation is undefined
            sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 4).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)

200:        sheet2.Cells(1, NumberOfColumns + 5).Value2 = "Nash-Sut."
            If SatirToplami1 = 0 Then GoTo 300 'Nash-Sut cannot be calculated when the row is empty
            SatirToplami1 = sheet2.Cells((RowNumber - 1) * 6 + 4, NumberOfColumns + 3).Value2
            SatirToplami2 = sheet2.Cells((RowNumber - 1) * 6 + 5, NumberOfColumns + 3).Value2
            sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 5).Value2 = 1 - SatirToplami1 / SatirToplami2

300:        sheet2.Cells(1, NumberOfColumns + 6).Value2 = "RMSE"
            If SatirToplami1 = 0 Then GoTo 310 'RMSE cannot be calculated when the row is empty
            RMSE = Math.Sqrt(SatirToplami1 / excel.WorksheetFunction.CountA(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 4, 3), sheet2.Cells((RowNumber - 1) * 6 + 4, NumberOfColumns + 1))))
            sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 6).Value2 = RMSE

310:        sheet2.Cells(1, NumberOfColumns + 7).Value2 = "NRMSE"
            RowMax = excel.WorksheetFunction.Max(CorrelationRange1)
            RowMin = excel.WorksheetFunction.Min(CorrelationRange1)
            If RowMax = Nothing Then GoTo 400 'NRMSE cannot be calculated when the row is empty
            sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 7).Value2 = RMSE / (RowMax - RowMin)

400:        'Mean Absolute Scaled Error:
            Fark3 = 0
            For j = 2 To NumberOfColumns - 1
                Fark3 = Fark3 + Math.Abs(ObservedData(RowNumber + 1, j + 1) - ObservedData(RowNumber + 1, j))
            Next
            If RowMax = Nothing Then GoTo 500 'MASE cannot be calculated when all observations are qual or the line is empty
            Fark3 = Fark3 * (NumberOfColumns - 1) / (NumberOfColumns - 2)
            Fark3Toplami = excel.WorksheetFunction.Sum(sheet2.Range(sheet2.Cells((RowNumber - 1) * 6 + 6, 3), sheet2.Cells((RowNumber - 1) * 6 + 6, NumberOfColumns + 1)))
            sheet2.Cells(1, NumberOfColumns + 8).Value2 = "MASE"
            sheet2.Cells((RowNumber - 1) * 6 + 3, NumberOfColumns + 8).Value2 = Fark3Toplami / Fark3

            sheet2.Columns("A:U").EntireColumn.AutoFit()

500:        'Relocation of the removed observed values to the rows
            For ColumnNumber = 1 To NumberOfColumns - 1
                nan(RowNumber, ColumnNumber) = RemovedData(ColumnNumber)
                nanWide(RowNumber + 3, ColumnNumber + 3) = RemovedData(ColumnNumber)
                If ColumnNumber < 4 Then nanWide(RowNumber + 2, ColumnNumber + NumberOfColumns + 2) = RemovedData(ColumnNumber)
                If ColumnNumber > 12 Then nanWide(RowNumber + 4, ColumnNumber - NumberOfColumns + 1) = RemovedData(ColumnNumber)
            Next
        Next ' RowNumber loop

        'Calculating and writing the statistics to the Cumulative Statistics and Graphs file
        For j = 0 To 5
            sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 4 + j, 2).Interior.Color = RGB(170 - j * 10, 250 - j * 10, 170 - j * 10)        'Green Tones
            sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 4 + j, 1).Interior.Color = RGB(200 - j * 10, 250 - j * 10, 170 - j * 10)        'Green Tones
        Next
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 1).Value2 = "Average"
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 1).Interior.Color = RGB(190, 140, 150)                                                        'Dark Lilac  
        Ortalama1 = excel.WorksheetFunction.Average(sheet4.Range(sheet4.Cells(2, 2), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 2)))
        Ortalama2 = excel.WorksheetFunction.Average(sheet4.Range(sheet4.Cells(2, 3), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 3)))
        If Ortalama1 = Nothing Then Ortalama1 = 0
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 2).Value2 = Ortalama1
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 3).Value2 = Ortalama2

        For i = 1 To NumberOfRows - 1
            For j = 1 To NumberOfColumns - 1
                If ObservedData(i + 1, j + 1) = 0 And ObservedData(i + 1, j + 1) IsNot Nothing Then GoTo 510
                If ObservedData(i + 1, j + 1) = Nothing Then GoTo 515
510:            Fark1 = ObservedData(i + 1, j + 1) - BestCorrelatedEstimation(i, j)
                sheet4.Cells((i - 1) * (NumberOfColumns - 1) + j + 1, 4).Value2 = Math.Pow(Fark1, 2)
                Fark2 = ObservedData(i + 1, j + 1) - Ortalama1
                sheet4.Cells((i - 1) * (NumberOfColumns - 1) + j + 1, 5).Value2 = Math.Pow(Fark2, 2)
                sheet4.Cells((i - 1) * (NumberOfColumns - 1) + j + 1, 6).Value2 = Math.Abs(Fark1)
515:        Next
        Next

        Ortalama1 = excel.WorksheetFunction.Average(sheet4.Range(sheet4.Cells(2, 4), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 4)))
        Ortalama2 = excel.WorksheetFunction.Average(sheet4.Range(sheet4.Cells(2, 5), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 5)))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 4).Value2 = Ortalama1
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 5).Value2 = Ortalama2

        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 1).Value2 = "Sum"
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 1).Interior.Color = RGB(200, 150, 160)                                                        'Daha Dark Lilac
        SatirToplami1 = excel.WorksheetFunction.Sum(sheet4.Range(sheet4.Cells(2, 2), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 2)))
        SatirToplami2 = excel.WorksheetFunction.Sum(sheet4.Range(sheet4.Cells(2, 3), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 3)))
        If SatirToplami1 = Nothing Then SatirToplami1 = 0
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 2).Value2 = SatirToplami1
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 3).Value2 = SatirToplami2

        SatirToplami1 = excel.WorksheetFunction.Sum(sheet4.Range(sheet4.Cells(2, 4), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 4)))
        SatirToplami2 = excel.WorksheetFunction.Sum(sheet4.Range(sheet4.Cells(2, 5), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 5)))
        If SatirToplami1 = Nothing Then SatirToplami1 = 0
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 4).Value2 = SatirToplami1
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 5).Value2 = SatirToplami2

        SatirToplami1 = excel.WorksheetFunction.Sum(sheet4.Range(sheet4.Cells(2, 6), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 6)))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 6).Value2 = SatirToplami1

        SixthRowAverage = excel.WorksheetFunction.Average(sheet4.Range(sheet4.Cells(2, 6), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 6)))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 2, 6).Value2 = SixthRowAverage

        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 4, 1).Value2 = "Corr.:"
        CorrelationRange1 = sheet4.Range(sheet4.Cells(2, 2), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 2))
        CorrelationRange2 = sheet4.Range(sheet4.Cells(2, 3), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 3))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 4, 2).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)

        CorrelationRange2 = sheet4.Range(sheet4.Cells(2, 7), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 7))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 4, 7).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)

        CorrelationRange2 = sheet4.Range(sheet4.Cells(2, 8), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 8))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 4, 8).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)

        CorrelationRange2 = sheet4.Range(sheet4.Cells(2, 9), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 9))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 4, 9).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)

        CorrelationRange2 = sheet4.Range(sheet4.Cells(2, 10), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 10))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 4, 10).Value2 = excel.WorksheetFunction.Correl(CorrelationRange1, CorrelationRange2)

        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 5, 1).Value2 = "Nash-Sut."
        SatirToplami1 = sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 4).Value2
        SatirToplami2 = sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 5).Value2
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 5, 2).Value2 = 1 - SatirToplami1 / SatirToplami2

600:    sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 6, 1).Value2 = "RMSE"
        RMSE = Math.Sqrt(SatirToplami1 / excel.WorksheetFunction.CountA(sheet4.Range(sheet4.Cells(2, 4), sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 1, 4))))
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 6, 2).Value2 = RMSE

        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 7, 1).Value2 = "NRMSE"
        RowMax = excel.WorksheetFunction.Max(CorrelationRange1)
        RowMin = excel.WorksheetFunction.Min(CorrelationRange1)
        If RowMax = Nothing Then GoTo 700 'NRMSE cannot be calculated when the row is empty
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 7, 2).Value2 = RMSE / (RowMax - RowMin)

700:    'Mean Asolute Scaled Error:
        Fark3 = 0
        For j = 3 To (NumberOfRows - 1) * (NumberOfColumns - 1) + 1
            Fark3 = Fark3 + Math.Abs(sheet4.Cells(j, 2).Value2 - sheet4.Cells(j - 1, 2).Value2)
        Next
        If RowMax = Nothing Then GoTo 800 'MASE cannot be calculated when all observations are qual or the line is empty
        Fark3 = Fark3 * ((NumberOfRows - 1) * (NumberOfColumns - 1) - 1) / ((NumberOfRows - 1) * (NumberOfColumns - 1) - 2)
        Fark3Toplami = sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 3, 6).Value2
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 8, 1).Value2 = "MASE"
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 8, 2).Value2 = Fark3Toplami / Fark3

        NumberOfCells = (NumberOfRows - 1) * (NumberOfColumns - 1)
        NumberOfMissingData = NumberOfCells - ElemanSayisi
        MissingDataRate = NumberOfMissingData / NumberOfCells
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 9, 1).Value2 = "MDR"
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 9, 2).NumberFormat = "0.000"
        sheet4.Cells((NumberOfRows - 1) * (NumberOfColumns - 1) + 9, 2).Value2 = MissingDataRate

        sheet4.Columns("A:F").EntireColumn.AutoFit()

800:    w.Save()
        w2.Save()
        w3.Save()
        w4.Save()
        w.Close()
        w2.Close()
        w3.Close()
        w4.Close()
    End Sub
End Module
