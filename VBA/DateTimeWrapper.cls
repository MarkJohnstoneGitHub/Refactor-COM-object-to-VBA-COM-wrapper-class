VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an instant in time, typically expressed as a date and time of day."
'@Folder("DotNet.System")
'@PredeclaredId
'@ModuleDescription("Represents an instant in time, typically expressed as a date and time of day.")

'https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

Option Explicit

Private Type TDateTimeWrapper
   DotNetLibDateTime As DotNetLib.DateTime
End Type

Private this As TDateTimeWrapper

Private Sub Class_Initialize()
   Set this.DotNetLibDateTime = New DotNetLib.DateTime
End Sub

Private Sub Class_Terminate()
   Set this.DotNetLibDateTime = Nothing
End Sub

Friend Property Get ComObject() As DotNetLib.DateTime
   Set ComObject = this.DotNetLibDateTime
End Property

Friend Property Set ComObject(ByVal objDateTime As DotNetLib.DateTime)
   Set this.DotNetLibDateTime = objDateTime
End Property

Friend Property Get Self() As DateTimeWrapper
   Set Self = Me
End Property

'@Description("Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time.")
Public Function CreateFromTicks(ByVal Ticks As LongLong, Optional ByVal kind As DotNetLib.DateTimeKind = DateTimeKind_Unspecified) As DateTimeWrapper
Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.CreateFromTicks(Ticks, kind)
      Set CreateFromTicks = .Self
   End With
End Function

'@Description("Initializes a new instance of the DateTime structure to the specified year, month, and day.")
Public Function CreateFromDate(ByVal year As Long, ByVal month As Long, ByVal Day As Long) As DateTimeWrapper
Attribute CreateFromDate.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, and day."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.CreateFromDate(year, month, Day)
      Set CreateFromDate = .Self
   End With
End Function

'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond.")
Public Function CreateFromDateTime(ByVal year As Long, ByVal month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, Optional ByVal Millisecond As Long = 0) As DateTimeWrapper
Attribute CreateFromDateTime.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.CreateFromDateTime(year, month, Day, Hour, Minute, Second, Millisecond)
      Set CreateFromDateTime = .Self
   End With
End Function

'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and Coordinated Universal Time (UTC) or local time.")
Public Function CreateFromDateTimeKind(ByVal year As Long, ByVal month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal kind As DotNetLib.DateTimeKind) As DateTimeWrapper
Attribute CreateFromDateTimeKind.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and Coordinated Universal Time (UTC) or local time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.CreateFromDateTimeKind(year, month, Day, Hour, Minute, Second, kind)
      Set CreateFromDateTimeKind = .Self
   End With
End Function

'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time.")
Public Function CreateFromDateTimeKind2(ByVal year As Long, ByVal month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, ByVal kind As DotNetLib.DateTimeKind) As DateTimeWrapper
Attribute CreateFromDateTimeKind2.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.CreateFromDateTimeKind2(year, month, Day, Hour, Minute, Second, Millisecond, kind)
      Set CreateFromDateTimeKind2 = .Self
   End With
End Function

'@Description("Represents the largest possible value of DateTime. This field is read-only.")
Public Property Get MaxValue() As DateTimeWrapper
Attribute MaxValue.VB_Description = "Represents the largest possible value of DateTime. This field is read-only."
    Set MaxValue = DateTimeStatic.MaxValue
End Property

'@Description("Represents the smallest possible value of DateTime. This field is read-only.")
Public Property Get MinValue() As DateTimeWrapper
Attribute MinValue.VB_Description = "Represents the smallest possible value of DateTime. This field is read-only."
    Set MinValue = DateTimeStatic.MinValue
End Property

'@Description("Gets the date component of this instance.")
Public Property Get DateComponent() As DateTimeWrapper
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Date()
      Set DateComponent = .Self
   End With
End Property

'@Description("Gets the day of the month represented by this instance.")
Public Property Get Day() As Long
Attribute Day.VB_Description = "Gets the day of the month represented by this instance."
   Day = this.DotNetLibDateTime.Day()
End Property

'@Description("Gets the day of the week represented by this instance.")
Public Property Get DayOfWeek() As DotNetLib.DayOfWeek
Attribute DayOfWeek.VB_Description = "Gets the day of the week represented by this instance."
   DayOfWeek = this.DotNetLibDateTime.DayOfWeek()
End Property

'@Description("Gets the day of the year represented by this instance.")
Public Property Get DayOfYear() As Long
Attribute DayOfYear.VB_Description = "Gets the day of the year represented by this instance."
   DayOfYear = this.DotNetLibDateTime.DayOfYear()
End Property

'@Description("Gets the hour component of the date represented by this instance.")
Public Property Get Hour() As Long
Attribute Hour.VB_Description = "Gets the hour component of the date represented by this instance."
   Hour = this.DotNetLibDateTime.Hour()
End Property

'@Description("Gets a value that indicates whether the time represented by this instance is based on local time, Coordinated Universal Time (UTC), or neither.")
Public Property Get kind() As DotNetLib.DateTimeKind
Attribute kind.VB_Description = "Gets a value that indicates whether the time represented by this instance is based on local time, Coordinated Universal Time (UTC), or neither."
   kind = this.DotNetLibDateTime.kind()
End Property

'@Description("Gets the milliseconds component of the date represented by this instance.")
Public Property Get Millisecond() As Long
Attribute Millisecond.VB_Description = "Gets the milliseconds component of the date represented by this instance."
   Millisecond = this.DotNetLibDateTime.Millisecond()
End Property

'@Description("Gets the minute component of the date represented by this instance.")
Public Property Get Minute() As Long
Attribute Minute.VB_Description = "Gets the minute component of the date represented by this instance."
   Minute = this.DotNetLibDateTime.Minute()
End Property

'@Description("Gets the month component of the date represented by this instance.")
Public Property Get month() As Long
Attribute month.VB_Description = "Gets the month component of the date represented by this instance."
   month = this.DotNetLibDateTime.month()
End Property

'@Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time.")
Public Property Get Now() As DateTimeWrapper
Attribute Now.VB_Description = "Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Now()
      Set Now = .Self
   End With
End Property

'@Description("Gets the seconds component of the date represented by this instance.")
Public Property Get Second() As Long
Attribute Second.VB_Description = "Gets the seconds component of the date represented by this instance."
   Second = this.DotNetLibDateTime.Second()
End Property

'@Description("Gets the number of ticks that represent the date and time of this instance.")
Public Property Get Ticks() As LongLong
Attribute Ticks.VB_Description = "Gets the number of ticks that represent the date and time of this instance."
   Ticks = this.DotNetLibDateTime.Ticks()
End Property

'@Description("Gets the time of day for this instance.")
Public Property Get TimeOfDay() As DotNetLib.TimeSpan
Attribute TimeOfDay.VB_Description = "Gets the time of day for this instance."
   Set TimeOfDay = this.DotNetLibDateTime.TimeOfDay()
End Property

'@Description("Gets the current date.")
Public Property Get Today() As DateTimeWrapper
Attribute Today.VB_Description = "Gets the current date."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Today()
      Set Today = .Self
   End With
End Property

'@Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC).")
Public Property Get UtcNow() As DateTimeWrapper
Attribute UtcNow.VB_Description = "Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC)."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.UtcNow()
      Set UtcNow = .Self
   End With
End Property

'@Description("Gets the year component of the date represented by this instance.")
Public Property Get year() As Long
Attribute year.VB_Description = "Gets the year component of the date represented by this instance."
   year = this.DotNetLibDateTime.year()
End Property

'@Description("Returns a new DateTime that adds the value of the specified TimeSpan to the value of this instance.")
Public Function Add(ByRef value As DotNetLib.TimeSpan) As DateTimeWrapper
Attribute Add.VB_Description = "Returns a new DateTime that adds the value of the specified TimeSpan to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Add(value)
      Set Add = .Self
   End With
End Function

'@Description("Returns a new DateTime that adds the specified number of days to the value of this instance.")
Public Function AddDays(ByVal value As Double) As DateTimeWrapper
Attribute AddDays.VB_Description = "Returns a new DateTime that adds the specified number of days to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.AddDays(value)
      Set AddDays = .Self
   End With
End Function

'@Description("Returns a new DateTime that adds the specified number of hours to the value of this instance.")
Public Function AddHours(ByVal value As Double) As DateTimeWrapper
Attribute AddHours.VB_Description = "Returns a new DateTime that adds the specified number of hours to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.AddHours(value)
      Set AddHours = .Self
   End With
End Function

'@Description("Returns a new DateTime that adds the specified number of milliseconds to the value of this instance.")
Public Function AddMilliseconds(ByVal value As Double) As DateTimeWrapper
Attribute AddMilliseconds.VB_Description = "Returns a new DateTime that adds the specified number of milliseconds to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.AddMilliseconds(value)
      Set AddMilliseconds = .Self
   End With
End Function

'@Description("Returns a new DateTime that adds the specified number of minutes to the value of this instance.")
Public Function AddMinutes(ByVal value As Double) As DateTimeWrapper
Attribute AddMinutes.VB_Description = "Returns a new DateTime that adds the specified number of minutes to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.AddMinutes(value)
      Set AddMinutes = .Self
   End With
End Function

'@Description("Returns a new DateTime that adds the specified number of months to the value of this instance.")
Public Function AddMonths(ByVal months As Long) As DateTimeWrapper
Attribute AddMonths.VB_Description = "Returns a new DateTime that adds the specified number of months to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.AddMonths(months)
      Set AddMonths = .Self
   End With
End Function

'@Description("Returns a new DateTime that adds the specified number of seconds to the value of this instance.")
Public Function AddSeconds(ByVal value As Double) As DateTimeWrapper
Attribute AddSeconds.VB_Description = "Returns a new DateTime that adds the specified number of seconds to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.AddSeconds(value)
      Set AddSeconds = .Self
   End With
End Function

'@Description("Returns a new DateTime that adds the specified number of ticks to the value of this instance.")
Public Function AddTicks(ByVal value As LongLong) As DateTimeWrapper
Attribute AddTicks.VB_Description = "Returns a new DateTime that adds the specified number of ticks to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.AddTicks(value)
      Set AddTicks = .Self
   End With
End Function

'@Description("Returns a new DateTime that adds the specified number of years to the value of this instance.")
Public Function AddYears(ByVal value As Long) As DateTimeWrapper
Attribute AddYears.VB_Description = "Returns a new DateTime that adds the specified number of years to the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.AddYears(value)
      Set AddYears = .Self
   End With
End Function

'@Description("Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance.")
Public Function Compare(ByRef t1 As DateTimeWrapper, ByRef t2 As DateTimeWrapper) As Long
Attribute Compare.VB_Description = "Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance."
   Compare = this.DotNetLibDateTime.Compare(t1.ComObject, t2.ComObject)
End Function

'@Description("Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.")
Public Function CompareTo(ByRef t As DateTimeWrapper) As Long
Attribute CompareTo.VB_Description = "Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value."
   CompareTo = this.DotNetLibDateTime.CompareTo(t.ComObject)
End Function

'@Description("Returns the number of days in the specified month and year.")
Public Function DaysInMonth(ByVal year As Long, ByVal month As Long) As Long
Attribute DaysInMonth.VB_Description = "Returns the number of days in the specified month and year."
   DaysInMonth = this.DotNetLibDateTime.DaysInMonth(year, month)
End Function

'@Description("Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance.")
Public Function Equals(ByRef value As DateTimeWrapper) As Boolean
Attribute Equals.VB_Description = "Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance."
   Equals = this.DotNetLibDateTime.Equals(value.ComObject)
End Function

'@Description("Returns a value indicating whether this instance is equal to a specified object.")
Public Function Equals2(ByVal value As Variant) As Boolean
Attribute Equals2.VB_Description = "Returns a value indicating whether this instance is equal to a specified object."
   Equals2 = this.DotNetLibDateTime.Equals2(value)
End Function

'@Description("Returns a value indicating whether two DateTime instances have the same date and time value.")
Public Function Equals3(ByRef t1 As DateTimeWrapper, ByRef t2 As DateTimeWrapper) As Boolean
Attribute Equals3.VB_Description = "Returns a value indicating whether two DateTime instances have the same date and time value."
   Equals3 = this.DotNetLibDateTime.Equals3(t1.ComObject, t2.ComObject)
End Function

'@Description("Deserializes a 64-bit binary value and recreates an original serialized DateTime object.")
Public Function FromBinary(ByVal dateData As LongLong) As DateTimeWrapper
Attribute FromBinary.VB_Description = "Deserializes a 64-bit binary value and recreates an original serialized DateTime object."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.FromBinary(dateData)
      Set FromBinary = .Self
   End With
End Function

'@Description("Converts the specified Windows file time to an equivalent local time.")
Public Function FromFileTime(ByVal fileTime As LongLong) As DateTimeWrapper
Attribute FromFileTime.VB_Description = "Converts the specified Windows file time to an equivalent local time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.FromFileTime(fileTime)
      Set FromFileTime = .Self
   End With
End Function

'@Description("Converts the specified Windows file time to an equivalent UTC time.")
Public Function FromFileTimeUtc(ByVal fileTime As LongLong) As DateTimeWrapper
Attribute FromFileTimeUtc.VB_Description = "Converts the specified Windows file time to an equivalent UTC time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.FromFileTimeUtc(fileTime)
      Set FromFileTimeUtc = .Self
   End With
End Function

'@Description("Returns a DateTime equivalent to the specified OLE Automation Date.")
Public Function FromOADate(ByVal d As Double) As DateTimeWrapper
Attribute FromOADate.VB_Description = "Returns a DateTime equivalent to the specified OLE Automation Date."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.FromOADate(d)
      Set FromOADate = .Self
   End With
End Function

'@Description("Converts the value of this instance to all the string representations supported by the standard date and time format specifiers.")
Public Function GetDateTimeFormats() As String()
Attribute GetDateTimeFormats.VB_Description = "Converts the value of this instance to all the string representations supported by the standard date and time format specifiers."
   GetDateTimeFormats = this.DotNetLibDateTime.GetDateTimeFormats()
End Function

'@Description("Returns the hash code for this instance.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Returns the hash code for this instance."
   GetHashCode = this.DotNetLibDateTime.GetHashCode()
End Function

'@Description("Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone.")
Public Function IsDaylightSavingTime() As Boolean
Attribute IsDaylightSavingTime.VB_Description = "Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone."
   IsDaylightSavingTime = this.DotNetLibDateTime.IsDaylightSavingTime()
End Function

'@Description("Returns an indication whether the specified year is a leap year.")
Public Function IsLeapYear(ByVal year As Long) As Boolean
Attribute IsLeapYear.VB_Description = "Returns an indication whether the specified year is a leap year."
   IsLeapYear = this.DotNetLibDateTime.IsLeapYear(year)
End Function

'@Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")
Public Function Parse(ByVal s As String) As DateTimeWrapper
Attribute Parse.VB_Description = "Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Parse(s)
      Set Parse = .Self
   End With
End Function

'@Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")
Public Function Parse2(ByVal s As String, ByRef provider As IFormatProvider) As DateTimeWrapper
Attribute Parse2.VB_Description = "Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Parse2(s, provider)
      Set Parse2 = .Self
   End With
End Function

'@Description("Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values.")
Public Function Parse3(ByVal s As String, ByRef provider As IFormatProvider, ByVal styles As DateTimeStyles) As DateTimeWrapper
Attribute Parse3.VB_Description = "Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Parse3(s, provider, styles)
      Set Parse3 = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function ParseExact(ByVal s As String, ByVal format As String, ByRef provider As IFormatProvider) As DateTimeWrapper
Attribute ParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.ParseExact(s, format, provider)
      Set ParseExact = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown.")
Public Function ParseExact2(ByVal s As String, ByVal format As String, ByRef provider As IFormatProvider, ByVal style As DateTimeStyles) As DateTimeWrapper
Attribute ParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.ParseExact2(s, format, provider, style)
      Set ParseExact2 = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown.")
Public Function ParseExact3(ByVal s As String, ByRef formats() As String, ByRef provider As IFormatProvider, ByVal style As DateTimeStyles) As DateTimeWrapper
Attribute ParseExact3.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.ParseExact3(s, formats, provider, style)
      Set ParseExact3 = .Self
   End With
End Function

'@Description("Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value.")
Public Function SpecifyKind(ByRef value As DateTimeWrapper, ByVal kind As DotNetLib.DateTimeKind) As DateTimeWrapper
Attribute SpecifyKind.VB_Description = "Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.SpecifyKind(value.ComObject, kind)
      Set SpecifyKind = .Self
   End With
End Function

'@Description("Returns a new DateTime that subtracts the specified duration from the value of this instance.")
Public Function Subtract(ByRef value As DotNetLib.TimeSpan) As DateTimeWrapper
Attribute Subtract.VB_Description = "Returns a new DateTime that subtracts the specified duration from the value of this instance."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Subtract(value)
      Set Subtract = .Self
   End With
End Function

'@Description("Returns a new TimeSpan that subtracts the specified date and time from the value of this instance.")
Public Function Subtract2(ByRef value As DateTimeWrapper) As DotNetLib.TimeSpan
Attribute Subtract2.VB_Description = "Returns a new TimeSpan that subtracts the specified date and time from the value of this instance."
   Set Subtract2 = this.DotNetLibDateTime.Subtract2(value.ComObject)
End Function

'@Description("Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object.")
Public Function ToBinary() As LongLong
Attribute ToBinary.VB_Description = "Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object."
   ToBinary = this.DotNetLibDateTime.ToBinary()
End Function

'@Description("Converts the value of the current DateTime object to a Windows file time.")
Public Function ToFileTime() As LongLong
Attribute ToFileTime.VB_Description = "Converts the value of the current DateTime object to a Windows file time."
   ToFileTime = this.DotNetLibDateTime.ToFileTime()
End Function

'@Description("Converts the value of the current DateTime object to a Windows file time.")
Public Function ToFileTimeUtc() As LongLong
Attribute ToFileTimeUtc.VB_Description = "Converts the value of the current DateTime object to a Windows file time."
   ToFileTimeUtc = this.DotNetLibDateTime.ToFileTimeUtc()
End Function

'@Description("Converts the value of the current DateTime object to local time.")
Public Function ToLocalTime() As DateTimeWrapper
Attribute ToLocalTime.VB_Description = "Converts the value of the current DateTime object to local time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.ToLocalTime()
      Set ToLocalTime = .Self
   End With
End Function

'@Description("Converts the value of the current DateTime object to its equivalent long date string representation.")
Public Function ToLongDateString() As String
Attribute ToLongDateString.VB_Description = "Converts the value of the current DateTime object to its equivalent long date string representation."
   ToLongDateString = this.DotNetLibDateTime.ToLongDateString()
End Function

'@Description("Converts the value of the current DateTime object to its equivalent long time string representation.")
Public Function ToLongTimeString() As String
Attribute ToLongTimeString.VB_Description = "Converts the value of the current DateTime object to its equivalent long time string representation."
   ToLongTimeString = this.DotNetLibDateTime.ToLongTimeString()
End Function

'@Description("Converts the value of this instance to the equivalent OLE Automation date.")
Public Function ToOADate() As Double
Attribute ToOADate.VB_Description = "Converts the value of this instance to the equivalent OLE Automation date."
   ToOADate = this.DotNetLibDateTime.ToOADate()
End Function

'@Description("Converts the value of the current DateTime object to its equivalent short date string representation.")
Public Function ToShortDateString() As String
Attribute ToShortDateString.VB_Description = "Converts the value of the current DateTime object to its equivalent short date string representation."
   ToShortDateString = this.DotNetLibDateTime.ToShortDateString()
End Function

'@Description("Converts the value of the current DateTime object to its equivalent short time string representation.")
Public Function ToShortTimeString() As String
Attribute ToShortTimeString.VB_Description = "Converts the value of the current DateTime object to its equivalent short time string representation."
   ToShortTimeString = this.DotNetLibDateTime.ToShortTimeString()
End Function

'@DefaultMember
'@Description("Converts the value of the current DateTime object to its equivalent string representation using the formatting conventions of the current culture.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the formatting conventions of the current culture."
Attribute ToString.VB_UserMemId = 0
   ToString = this.DotNetLibDateTime.ToString()
End Property

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified format and the formatting conventions of the current culture.")
Public Function ToString2(ByVal format As String) As String
Attribute ToString2.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified format and the formatting conventions of the current culture."
   ToString2 = this.DotNetLibDateTime.ToString2(format)
End Function

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified culture-specific format information.")
Public Function ToString3(ByRef provider As IFormatProvider) As String
Attribute ToString3.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified culture-specific format information."
   ToString3 = this.DotNetLibDateTime.ToString3(provider)
End Function

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified format and culture-specific format information.")
Public Function ToString4(ByVal format As String, ByRef provider As IFormatProvider) As String
Attribute ToString4.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified format and culture-specific format information."
   ToString4 = this.DotNetLibDateTime.ToString4(format, provider)
End Function

'@Description("Converts the value of the current DateTime object to Coordinated Universal Time (UTC).")
Public Function ToUniversalTime() As DateTimeWrapper
Attribute ToUniversalTime.VB_Description = "Converts the value of the current DateTime object to Coordinated Universal Time (UTC)."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.ToUniversalTime()
      Set ToUniversalTime = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTime equivalent and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse(ByVal s As String, ByRef result As DateTimeWrapper) As Boolean
Attribute TryParse.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent and returns a value that indicates whether the conversion succeeded."
   TryParse = this.DotNetLibDateTime.TryParse(s, result.ComObject)
End Function

'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified culture-specific format information and formatting style, and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse2(ByVal s As String, ByRef provider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DateTimeWrapper) As Boolean
Attribute TryParse2.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified culture-specific format information and formatting style, and returns a value that indicates whether the conversion succeeded."
   TryParse2 = this.DotNetLibDateTime.TryParse2(s, provider, styles, result.ComObject)
End Function

'@Description("Adds a specified time interval to a specified date and time, yielding a new date and time.")
Public Function Addition(ByRef d As DateTimeWrapper, ByRef t As DotNetLib.TimeSpan) As DateTimeWrapper
Attribute Addition.VB_Description = "Adds a specified time interval to a specified date and time, yielding a new date and time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Addition(d.ComObject, t)
      Set Addition = .Self
   End With
End Function

'@Description("Determines whether two specified instances of DateTime are equal.")
Public Function Equality(ByRef d1 As DateTimeWrapper, ByRef d2 As DateTimeWrapper) As Boolean
Attribute Equality.VB_Description = "Determines whether two specified instances of DateTime are equal."
   Equality = this.DotNetLibDateTime.Equality(d1.ComObject, d2.ComObject)
End Function

'@Description("Determines whether one specified DateTime is later than another specified DateTime.")
Public Function GreaterThan(ByRef t1 As DateTimeWrapper, ByRef t2 As DateTimeWrapper) As Boolean
Attribute GreaterThan.VB_Description = "Determines whether one specified DateTime is later than another specified DateTime."
   GreaterThan = this.DotNetLibDateTime.GreaterThan(t1.ComObject, t2.ComObject)
End Function

'@Description("Determines whether one specified DateTime represents a date and time that is the same as or later than another specified DateTime.")
Public Function GreaterThanOrEqual(ByRef t1 As DateTimeWrapper, ByRef t2 As DateTimeWrapper) As Boolean
Attribute GreaterThanOrEqual.VB_Description = "Determines whether one specified DateTime represents a date and time that is the same as or later than another specified DateTime."
   GreaterThanOrEqual = this.DotNetLibDateTime.GreaterThanOrEqual(t1.ComObject, t2.ComObject)
End Function

'@Description("Determines whether two specified instances of DateTime are not equal.")
Public Function Inequality(ByRef t1 As DateTimeWrapper, ByRef t2 As DateTimeWrapper) As Boolean
Attribute Inequality.VB_Description = "Determines whether two specified instances of DateTime are not equal."
   Inequality = this.DotNetLibDateTime.Inequality(t1.ComObject, t2.ComObject)
End Function

'@Description("Determines whether one specified DateTime is earlier than another specified DateTime.")
Public Function LessThan(ByRef t1 As DateTimeWrapper, ByRef t2 As DateTimeWrapper) As Boolean
Attribute LessThan.VB_Description = "Determines whether one specified DateTime is earlier than another specified DateTime."
   LessThan = this.DotNetLibDateTime.LessThan(t1.ComObject, t2.ComObject)
End Function

'@Description("Determines whether one specified DateTime represents a date and time that is the same as or earlier than another specified DateTime.")
Public Function LessThanOrEqual(ByRef t1 As DateTimeWrapper, ByRef t2 As DateTimeWrapper) As Boolean
Attribute LessThanOrEqual.VB_Description = "Determines whether one specified DateTime represents a date and time that is the same as or earlier than another specified DateTime."
   LessThanOrEqual = this.DotNetLibDateTime.LessThanOrEqual(t1.ComObject, t2.ComObject)
End Function

'@Description("Subtracts a specified date and time from another specified date and time and returns a time interval.")
Public Function Subtraction(ByRef d1 As DateTimeWrapper, ByRef d2 As DateTimeWrapper) As DotNetLib.TimeSpan
Attribute Subtraction.VB_Description = "Subtracts a specified date and time from another specified date and time and returns a time interval."
   Set Subtraction = this.DotNetLibDateTime.Subtraction(d1.ComObject, d2.ComObject)
End Function

'@Description("Subtracts a specified time interval from a specified date and time and returns a new date and time.")
Public Function Subtraction2(ByRef d As DateTimeWrapper, ByRef t As DotNetLib.TimeSpan) As DateTimeWrapper
Attribute Subtraction2.VB_Description = "Subtracts a specified time interval from a specified date and time and returns a new date and time."
   With New DateTimeWrapper
      Set .ComObject = this.DotNetLibDateTime.Subtraction2(d.ComObject, t)
      Set Subtraction2 = .Self
   End With
End Function

