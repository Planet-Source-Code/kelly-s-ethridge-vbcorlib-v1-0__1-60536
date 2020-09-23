Attribute VB_Name = "modMain"
Option Explicit

'Note on Project properties
'Standard Exe's do not export a type library so the introspection using typelib32
'Thats Why its an ActiveXExe
'To make this work like a standard exe at startup, we do this....
'Startup Object is 'Sub Main'
'Start Mode is 'Standalone'
Sub Main()
   'AddTest New BasicTests
    AddTests

   If Command() = "/nogui" Then
      'use this call to run the tests without hooking into the ocx
      'errors will be logged to a file named Errors.txt and a report to Report.txt
      'at the moment these files go into the folder where the ocx is registered.  This is a TODO change ;)
      RunTestsWithoutGui
   Else
      frmSimplyVBUnitRunner.Show
   End If
End Sub

Private Sub AddTests()
#Const ALLTESTS = 1


' StreamReader Determine Encoding tests
#If False Or ALLTESTS Then
    AddTest New TestDetermineEncoding
#End If
' BinaryReader tests
#If False Or ALLTESTS Then
    AddTest New TestBinaryReader
#End If
' BinaryWriter tests
#If False Or ALLTESTS Then
    AddTest New TestBinaryWriter
#End If
' FileInfo tests
#If False Or ALLTESTS Then
    AddTest New TestFileInfo
#End If
' File tests
#If False Or ALLTESTS Then
    AddTest New TestFile
#End If
' StreamReader tests
#If False Or ALLTESTS Then
    AddTest New TestStreamReader
#End If
' StreamWriter tests
#If False Or ALLTESTS Then
    AddTest New TestStreamWriter
    AddTest New TestStreamWriterWithMem
    AddTest New TestSWWithMemAutoFlush
#End If
' Directory tests
#If False Or ALLTESTS Then
    AddTest New TestDirectory
#End If
' DirectoryInfo tests
#If False Or ALLTESTS Then
    AddTest New TestDirectoryInfo
#End If
' StringReader tests
#If False Or ALLTESTS Then
    AddTest New TestStringReader
#End If
' StringWriter tests
#If False Or ALLTESTS Then
    AddTest New TestStringWriter
#End If
' UnicodeEncoding with big-endian order tests
#If False Or ALLTESTS Then
    AddTest New TestUnicodeEncodingBig
#End If
' UnicodeEncoding tests
#If False Or ALLTESTS Then
    AddTest New TestUnicodeEncoding
#End If
' FileStream tests
#If False Or ALLTESTS Then
    AddTest New TestFileStreamWrite
    AddTest New TestFileStreamSmallBuffer
    AddTest New TestFileStream
#End If
' MemoryStream tests
#If False Or ALLTESTS Then
    AddTest New TestUserMemoryStream
    AddTest New TestMemoryStream
#End If
' UTF7Encoding tests
#If False Or ALLTESTS Then
    AddTest New TestUTF7GetChars
    AddTest New TestUTF7GetCharCount
    AddTest New TestUTF7GetBytes
    AddTest New TestUTF7GetByteCount
#End If
' UTF8Encoding tests
#If False Or ALLTESTS Then
    AddTest New TestUTF8GetChars
    AddTest New TestUTF8GetCharCount
    AddTest New TestUTF8Encoding
    AddTest New TestUTF8GetByteCount
#End If
' Path tests
#If False Or ALLTESTS Then
    AddTest New TestPath
#End If
' This only displays outputs from the Environment class.
#If False Or ALLTESTS Then
    AddTest New TestEnvironment
#End If
' TimeZone tests
#If False Or ALLTESTS Then
    AddTest New TestTimeZone
#End If
' DateTimeFormatInfo Invariant tests
#If False Or ALLTESTS Then
    AddTest New TestDateTimeFormatInfoInv
#End If
' CultureInfo tests
#If False Or ALLTESTS Then
    AddTest New TestCultureInfo
#End If
' MappedFile tests
#If False Or ALLTESTS Then
    AddTest New TestMappedFile
#End If
' FileNotFoundException tests
#If False Or ALLTESTS Then
    AddTest New TestFileNotFoundException
#End If
' cDateTime tests
#If False Or ALLTESTS Then
    AddTest New TestcDateTime
#End If
' TimeSpan tests
#If False Or ALLTESTS Then
    AddTest New TestTimeSpan
    AddTest New TestTimeSpan994394150ms
    AddTest New TestTimeSpanCreation
#End If
' Version tests
#If False Or ALLTESTS Then
    AddTest New TestVersion
#End If
' Random tests
#If False Or ALLTESTS Then
    AddTest New TestRandom
#End If
' BitConverter tests
#If False Or ALLTESTS Then
    AddTest New TestBitConverter
#End If
' NumberFormatInfo tests
#If False Or ALLTESTS Then
    AddTest New TestNumberFormatInfoInt
    AddTest New TestNumberFormatInfoFlt
    AddTest New TestNumberFormatInfoSng
#End If
' WeakReference tests
#If False Or ALLTESTS Then
    AddTest New TestWeakReference
#End If
' Hashtable tests
#If False Or ALLTESTS Then
    AddTest New TestHashTable
#End If
' Buffer tests
#If False Or ALLTESTS Then
    AddTest New TestBuffer
#End If
' BitArray tests
#If False Or ALLTESTS Then
    AddTest New TestBitArray
#End If
' SortedList tests
#If False Or ALLTESTS Then
    AddTest New TestSortedList
#End If
' DictionaryEntry
#If False Or ALLTESTS Then
    AddTest New TestDictionaryEntry
#End If
' Queue tests
#If False Or ALLTESTS Then
    AddTest New TestQueue
#End If
' Stack tests
#If False Or ALLTESTS Then
    AddTest New TestStack
#End If
' ArrayList tests
#If False Or ALLTESTS Then
    AddTest New TestArrayListExceptions
    AddTest New TestArrayListRange
    AddTest New TestArrayList10Items
    AddTest New TestArrayList
#End If
' cString tests
#If False Or ALLTESTS Then
    AddTest New TestcString
#End If
' StringBuilder tests
#If False Or ALLTESTS Then
    AddTest New TestStringBuilder
#End If
' Exception tests
#If False Or ALLTESTS Then
    AddTest New TestException
    AddTest New TestDefaultException
    AddTest New TestSystemException
    AddTest New TestDefaultSystemEx
    AddTest New TestArgumentException
    AddTest New TestDefaultArgumentEx
    AddTest New TestDefaultArgumentNull
    AddTest New TestArgumentNullException
    AddTest New TestArgumentOutOfRange
    AddTest New TestDefArgumentOutOfRange
#End If

' cArray tests
#If False Or ALLTESTS Then
    AddTest New TestcArray
    AddTest New TestInvalidCastException
    AddTest New TestDefInvalidCast
    AddTest New TestExceptionMethods
    AddTest New TestDefaultComparer
    AddTest New TestPosNumBinarySearch
    AddTest New TestMixNumBinarySearch
    AddTest New TestArraySort
    AddTest New TestArrayBinarySearch
    AddTest New TestArrayReverse
    AddTest New TestArrayIndexOf
    AddTest New TestArrayLastIndexOf
    AddTest New TestArrayCopy
    AddTest New TestArrayCreation
#End If
End Sub
