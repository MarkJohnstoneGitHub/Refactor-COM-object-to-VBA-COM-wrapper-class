VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTime"
Attribute VB_Description = "Represents an instant in time, typically expressed as a date and time of day."
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed  = False

'@PredeclaredId
'@ModuleDescription("Represents an instant in time, typically expressed as a date and time of day.")

Option Explicit

Private Type TDateTime
   DotNetLibDateTime As DotNetLib.DateTime
End Type

Private this As TDateTime

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
   Set ComObject = this.DotNetLibDateTime
End Property

Friend Property Get Self() As DateTime
   Set Self = Me
End Property

'@Description("Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time.")
Public Function CreateFromTicks(ByVal ticks As LongLong, Optional ByVal kind As DateTimeKind = DateTimeKind_Unspecified) As DateTime
Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time."
End Function

'@Description("Initializes a new instance of the DateTime structure to the specified year, month, and day.")
Public Function CreateFromDate(ByVal year As Long, ByVal month As Long, ByVal day As Long) As DateTime
Attribute CreateFromDate.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, and day."
End Function

'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond.")
Public Function CreateFromDateTime(ByVal year As Long, ByVal month As Long, ByVal day As Long, ByVal hour As Long, ByVal minute As Long, ByVal second As Long, Optional ByVal millisecond As Long = 0) As DateTime
Attribute CreateFromDateTime.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond."
End Function

'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time.")
Public Function CreateFromDateTimeKind(ByVal year As Long, ByVal month As Long, ByVal day As Long, ByVal hour As Long, ByVal minute As Long, ByVal second As Long, ByVal millisecond As Long, ByVal kind As DateTimeKind) As DateTime
Attribute CreateFromDateTimeKind.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time."
End Function

'@Description("Represents the largest possible value of DateTime. This field is read-only.")
Public Property Get MaxValue() As DateTime
Attribute MaxValue.VB_Description = "Represents the largest possible value of DateTime. This field is read-only."
End Property

'@Description("Represents the smallest possible value of DateTime. This field is read-only.")
Public Property Get MinValue() As DateTime
Attribute MinValue.VB_Description = "Represents the smallest possible value of DateTime. This field is read-only."
End Property

'@Description("Gets the date component of this instance.")
Public Property Get Date() As DateTime
Attribute Date.VB_Description = "Gets the date component of this instance."
End Property

'@Description("Gets the day of the month represented by this instance.")
Public Property Get Day() As Long
Attribute Day.VB_Description = "Gets the day of the month represented by this instance."
End Property

'@Description("Gets the day of the week represented by this instance.")
Public Property Get DayOfWeek() As DayOfWeek
Attribute DayOfWeek.VB_Description = "Gets the day of the week represented by this instance."
End Property

'@Description("Gets the day of the year represented by this instance.")
Public Property Get DayOfYear() As Long
Attribute DayOfYear.VB_Description = "Gets the day of the year represented by this instance."
End Property

'@Description("Gets the hour component of the date represented by this instance.")
Public Property Get Hour() As Long
Attribute Hour.VB_Description = "Gets the hour component of the date represented by this instance."
End Property

'@Description("Gets a value that indicates whether the time represented by this instance is based on local time, Coordinated Universal Time (UTC), or neither.")
Public Property Get Kind() As DateTimeKind
Attribute Kind.VB_Description = "Gets a value that indicates whether the time represented by this instance is based on local time, Coordinated Universal Time (UTC), or neither."
End Property

'@Description("Gets the milliseconds component of the date represented by this instance.")
Public Property Get Millisecond() As Long
Attribute Millisecond.VB_Description = "Gets the milliseconds component of the date represented by this instance."
End Property

'@Description("Gets the minute component of the date represented by this instance.")
Public Property Get Minute() As Long
Attribute Minute.VB_Description = "Gets the minute component of the date represented by this instance."
End Property

'@Description("Gets the month component of the date represented by this instance.")
Public Property Get Month() As Long
Attribute Month.VB_Description = "Gets the month component of the date represented by this instance."
End Property

'@Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time.")
Public Property Get Now() As DateTime
Attribute Now.VB_Description = "Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time."
End Property

'@Description("Gets the seconds component of the date represented by this instance.")
Public Property Get Second() As Long
Attribute Second.VB_Description = "Gets the seconds component of the date represented by this instance."
End Property

'@Description("Gets the number of ticks that represent the date and time of this instance.")
Public Property Get Ticks() As LongLong
Attribute Ticks.VB_Description = "Gets the number of ticks that represent the date and time of this instance."
End Property

'@Description("Gets the time of day for this instance.")
Public Property Get TimeOfDay() As ITimeSpan
Attribute TimeOfDay.VB_Description = "Gets the time of day for this instance."
End Property

'@Description("Gets the current date.")
Public Property Get Today() As DateTime
Attribute Today.VB_Description = "Gets the current date."
End Property

'@Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC).")
Public Property Get UtcNow() As DateTime
Attribute UtcNow.VB_Description = "Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC)."
End Property

'@Description("Gets the year component of the date represented by this instance.")
Public Property Get Year() As Long
Attribute Year.VB_Description = "Gets the year component of the date represented by this instance."
End Property

'@Description("Returns a new DateTime that adds the value of the specified TimeSpan to the value of this instance.")
Public Function Add(ByRef value As ITimeSpan) As DateTime
Attribute Add.VB_Description = "Returns a new DateTime that adds the value of the specified TimeSpan to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of days to the value of this instance.")
Public Function AddDays(ByVal value As Double) As DateTime
Attribute AddDays.VB_Description = "Returns a new DateTime that adds the specified number of days to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of hours to the value of this instance.")
Public Function AddHours(ByVal value As Double) As DateTime
Attribute AddHours.VB_Description = "Returns a new DateTime that adds the specified number of hours to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of milliseconds to the value of this instance.")
Public Function AddMilliseconds(ByVal value As Double) As DateTime
Attribute AddMilliseconds.VB_Description = "Returns a new DateTime that adds the specified number of milliseconds to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of minutes to the value of this instance.")
Public Function AddMinutes(ByVal value As Double) As DateTime
Attribute AddMinutes.VB_Description = "Returns a new DateTime that adds the specified number of minutes to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of months to the value of this instance.")
Public Function AddMonths(ByVal months As Long) As DateTime
Attribute AddMonths.VB_Description = "Returns a new DateTime that adds the specified number of months to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of seconds to the value of this instance.")
Public Function AddSeconds(ByVal value As Double) As DateTime
Attribute AddSeconds.VB_Description = "Returns a new DateTime that adds the specified number of seconds to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of ticks to the value of this instance.")
Public Function AddTicks(ByVal value As LongLong) As DateTime
Attribute AddTicks.VB_Description = "Returns a new DateTime that adds the specified number of ticks to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of years to the value of this instance.")
Public Function AddYears(ByVal value As Long) As DateTime
Attribute AddYears.VB_Description = "Returns a new DateTime that adds the specified number of years to the value of this instance."
End Function

'@Description("Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance.")
Public Function Compare(ByRef t1 As IDateTime, ByRef t2 As IDateTime) As Long
Attribute Compare.VB_Description = "Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance."
End Function

'@Description("Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.")
Public Function CompareTo(ByRef t As IDateTime) As Long
Attribute CompareTo.VB_Description = "Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value."
End Function

'@Description("Returns the number of days in the specified month and year.")
Public Function DaysInMonth(ByVal year As Long, ByVal month As Long) As Long
Attribute DaysInMonth.VB_Description = "Returns the number of days in the specified month and year."
End Function

'@Description("Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance.")
Public Function Equals(ByRef value As IDateTime) As Boolean
Attribute Equals.VB_Description = "Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance."
End Function

'@Description("Returns a value indicating whether this instance is equal to a specified object.")
Public Function Equals2(ByVal value As Variant) As Boolean
Attribute Equals2.VB_Description = "Returns a value indicating whether this instance is equal to a specified object."
End Function

'@Description("Returns a value indicating whether two DateTime instances have the same date and time value.")
Public Function Equals3(ByRef t1 As IDateTime, ByRef t2 As IDateTime) As Boolean
Attribute Equals3.VB_Description = "Returns a value indicating whether two DateTime instances have the same date and time value."
End Function

'@Description("Deserializes a 64-bit binary value and recreates an original serialized DateTime object.")
Public Function FromBinary(ByVal dateData As LongLong) As DateTime
Attribute FromBinary.VB_Description = "Deserializes a 64-bit binary value and recreates an original serialized DateTime object."
End Function

'@Description("Converts the specified Windows file time to an equivalent local time.")
Public Function FromFileTime(ByVal fileTime As LongLong) As DateTime
Attribute FromFileTime.VB_Description = "Converts the specified Windows file time to an equivalent local time."
End Function

'@Description("Converts the specified Windows file time to an equivalent UTC time.")
Public Function FromFileTimeUtc(ByVal fileTime As LongLong) As DateTime
Attribute FromFileTimeUtc.VB_Description = "Converts the specified Windows file time to an equivalent UTC time."
End Function

'@Description("Returns a DateTime equivalent to the specified OLE Automation Date.")
Public Function FromOADate(ByVal d As Double) As DateTime
Attribute FromOADate.VB_Description = "Returns a DateTime equivalent to the specified OLE Automation Date."
End Function

'@Description("Converts the value of this instance to all the string representations supported by the standard date and time format specifiers.")
Public Function GetDateTimeFormats() As String
Attribute GetDateTimeFormats.VB_Description = "Converts the value of this instance to all the string representations supported by the standard date and time format specifiers."
End Function

'@Description("Returns the hash code for this instance.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Returns the hash code for this instance."
End Function

'@Description("Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone.")
Public Function IsDaylightSavingTime() As Boolean
Attribute IsDaylightSavingTime.VB_Description = "Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone."
End Function

'@Description("Returns an indication whether the specified year is a leap year.")
Public Function IsLeapYear(ByVal year As Long) As Boolean
Attribute IsLeapYear.VB_Description = "Returns an indication whether the specified year is a leap year."
End Function

'@Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")
Public Function Parse(ByVal s As String) As DateTime
Attribute Parse.VB_Description = "Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture."
End Function

'@Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")
Public Function Parse2(ByVal s As String, ByRef provider As IFormatProvider) As DateTime
Attribute Parse2.VB_Description = "Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture."
End Function

'@Description("Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values.")
Public Function Parse3(ByVal s As String, ByRef provider As IFormatProvider, ByVal styles As DateTimeStyles) As DateTime
Attribute Parse3.VB_Description = "Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values."
End Function

'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function ParseExact(ByVal s As String, ByVal format As String, ByRef provider As IFormatProvider) As DateTime
Attribute ParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly."
End Function

'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown.")
Public Function ParseExact2(ByVal s As String, ByVal format As String, ByRef provider As IFormatProvider, ByVal style As DateTimeStyles) As DateTime
Attribute ParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown."
End Function

'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown.")
Public Function ParseExact3(ByVal s As String, ByVal formats As String, ByRef provider As IFormatProvider, ByVal style As DateTimeStyles) As DateTime
Attribute ParseExact3.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown."
End Function

'@Description("Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value.")
Public Function SpecifyKind(ByRef value As IDateTime, ByVal kind As DateTimeKind) As DateTime
Attribute SpecifyKind.VB_Description = "Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value."
End Function

'@Description("Returns a new DateTime that subtracts the specified duration from the value of this instance.")
Public Function Subtract(ByRef value As ITimeSpan) As DateTime
Attribute Subtract.VB_Description = "Returns a new DateTime that subtracts the specified duration from the value of this instance."
End Function

'@Description("Returns a new TimeSpan that subtracts the specified date and time from the value of this instance.")
Public Function Subtract2(ByRef value As IDateTime) As ITimeSpan
Attribute Subtract2.VB_Description = "Returns a new TimeSpan that subtracts the specified date and time from the value of this instance."
End Function

'@Description("Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object.")
Public Function ToBinary() As LongLong
Attribute ToBinary.VB_Description = "Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object."
End Function

'@Description("Converts the value of the current DateTime object to a Windows file time.")
Public Function ToFileTime() As LongLong
Attribute ToFileTime.VB_Description = "Converts the value of the current DateTime object to a Windows file time."
End Function

'@Description("Converts the value of the current DateTime object to a Windows file time.")
Public Function ToFileTimeUtc() As LongLong
Attribute ToFileTimeUtc.VB_Description = "Converts the value of the current DateTime object to a Windows file time."
End Function

'@Description("Converts the value of the current DateTime object to local time.")
Public Function ToLocalTime() As DateTime
Attribute ToLocalTime.VB_Description = "Converts the value of the current DateTime object to local time."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent long date string representation.")
Public Function ToLongDateString() As String
Attribute ToLongDateString.VB_Description = "Converts the value of the current DateTime object to its equivalent long date string representation."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent long time string representation.")
Public Function ToLongTimeString() As String
Attribute ToLongTimeString.VB_Description = "Converts the value of the current DateTime object to its equivalent long time string representation."
End Function

'@Description("Converts the value of this instance to the equivalent OLE Automation date.")
Public Function ToOADate() As Double
Attribute ToOADate.VB_Description = "Converts the value of this instance to the equivalent OLE Automation date."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent short date string representation.")
Public Function ToShortDateString() As String
Attribute ToShortDateString.VB_Description = "Converts the value of the current DateTime object to its equivalent short date string representation."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent short time string representation.")
Public Function ToShortTimeString() As String
Attribute ToShortTimeString.VB_Description = "Converts the value of the current DateTime object to its equivalent short time string representation."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the formatting conventions of the current culture.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the formatting conventions of the current culture."
End Property

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified format and the formatting conventions of the current culture.")
Public Function ToString2(ByVal format As String) As String
Attribute ToString2.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified format and the formatting conventions of the current culture."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified culture-specific format information.")
Public Function ToString3(ByRef provider As IFormatProvider) As String
Attribute ToString3.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified culture-specific format information."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified format and culture-specific format information.")
Public Function ToString4(ByVal format As String, ByRef provider As IFormatProvider) As String
Attribute ToString4.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified format and culture-specific format information."
End Function

Public Function ToUniversalTime() As DateTime
End Function

Public Function TryParse(ByVal s As String, ByRef result As IDateTime) As Boolean
End Function

Public Function Addition(ByRef d As IDateTime, ByRef t As ITimeSpan) As DateTime
End Function

Public Function Equality(ByRef d1 As IDateTime, ByRef d2 As IDateTime) As Boolean
End Function

Public Function GreaterThan(ByRef t1 As IDateTime, ByRef t2 As IDateTime) As Boolean
End Function

Public Function GreaterThanOrEqual(ByRef t1 As IDateTime, ByRef t2 As IDateTime) As Boolean
End Function

Public Function Inequality(ByRef t1 As IDateTime, ByRef t2 As IDateTime) As Boolean
End Function

Public Function LessThan(ByRef t1 As IDateTime, ByRef t2 As IDateTime) As Boolean
End Function

Public Function LessThanOrEqual(ByRef t1 As IDateTime, ByRef t2 As IDateTime) As Boolean
End Function

Public Function Subtraction(ByRef d1 As IDateTime, ByRef d2 As IDateTime) As ITimeSpan
End Function

Public Function Subtraction2(ByRef d As IDateTime, ByRef t As ITimeSpan) As DateTime
End Function
