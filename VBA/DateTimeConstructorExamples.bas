Attribute VB_Name = "DateTimeConstructorExamples"
'@Folder("Examples.DateTime")
Option Explicit


'@Description("This example demonstrates the DateTime(Int64) constructor.")
'https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int64)
Public Sub DateTimeCreateFromTicks()
    ' Instead of using the implicit, default "G" date and time format string, we
    ' use a custom format string that aligns the results and inserts leading zeroes.
    Dim format As String
    format = "MM/dd/yyyy hh:mm:ss tt"
    
    'Create a DateTime for the maximum date and time using ticks.
    Dim dt1 As DateTimeWrapper
    Set dt1 = DateTimeWrapper.CreateFromTicks(DateTimeWrapper.MaxValue.Ticks)
    
    'Create a DateTime for the minimum date and time using ticks.
    Dim dt2 As DateTimeWrapper
    Set dt2 = DateTimeWrapper.CreateFromTicks(DateTimeWrapper.MinValue.Ticks)
    
    'Create a custom DateTime for 7/28/1979 at 10:35:05 PM
    Dim pvtTicks As LongLong
    pvtTicks = DateTimeWrapper.CreateFromDateTime(1979, 7, 28, 22, 35, 5).Ticks
    Dim dt3 As DateTimeWrapper
    Set dt3 = DateTimeWrapper.CreateFromTicks(pvtTicks)
    
    Debug.Print "1) The maximum date and time is " & dt1.ToString2(format)
    Debug.Print "2) The minimum date and time is " & dt2.ToString2(format)
    Debug.Print "3) The custom  date and time is " & dt3.ToString2(format)
    Debug.Print "The custom date and time is created from " & pvtTicks & " ticks."
End Sub

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, DateTimeKind) constructor to instantiate a DateTime value.")
Public Sub DateTimeCreateFromDateTimeKind()
    Dim date1 As DateTimeWrapper
    Set date1 = DateTimeWrapper.CreateFromDateTimeKind(2010, 8, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)
    
    Debug.Print date1.ToString() & " " & DateTimeKindHelper.ToString(date1.kind)
    ' The example displays the following output, in this case for en-us culture:
    '      8/18/2010 4:32:00 PM Local
End Sub

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, Int32, DateTimeKind) constructor to instantiate a DateTime value.)"
Public Sub DateTimeCreateFromDateTimeKind2()
    Dim date1 As DateTimeWrapper
    Set date1 = DateTimeWrapper.CreateFromDateTimeKind2(2010, 8, 18, 16, 32, 18, 500, DateTimeKind.DateTimeKind_Local)
    
    Debug.Print date1.ToString2("M/dd/yyyy h:mm:ss.fff tt") & " " & DateTimeKindHelper.ToString(date1.kind)
    ' The example displays the following output, in this case for en-us culture:
    ' 8/18/2010 4:32:18.500 PM Local
End Sub

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value.")
Public Sub DateTimeCreateFromDateTime()
    Dim date1 As DateTimeWrapper
    Set date1 = DateTimeWrapper.CreateFromDateTime(2010, 8, 18, 16, 32, 18, 500)
    
    Debug.Print date1.ToString2("M/dd/yyyy h:mm:ss.fff tt")
    ' The example displays the following output, in this case for en-us culture:
    ' 8/18/2010 4:32:18.500 PM
End Sub


