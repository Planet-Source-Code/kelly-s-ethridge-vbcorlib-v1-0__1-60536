VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: FileInfo
'

''
' Provides methods for manipulating, setting and retrieving information for the specified file.
'
' @remarks The existence of the file is not verified until an action that requires
' accessing the file.
'
' @see Constructors
'
Option Explicit
Implements cObject
Implements FileSystemInfo

Private mFileName As String
Private mOriginalFileName As String
Private mIsDirty As Boolean
Private mFileAttributes As FileAttributes
Private mFileSize As Currency
Private mCreationTime As cDateTime
Private mLastAccessTime As cDateTime
Private mLastWriteTime As cDateTime



''
' Returns the attributes for the file (ReadOnly, Hidden, ...)
'
' @return The attributes of the file.
' @remarks The attributes are represented as bits. Each attributes has a specific bit.
' To test if a specific attribute exists, the attributes need to be ANDed with
' the individual attribute bits.
'
Public Property Get Attributes() As FileAttributes
    If mIsDirty Then Refresh
    Attributes = mFileAttributes
End Property

''
' Sets the attributes for the file (ReadOnly, Hiddin, ...)
'
' @param RHS The attributes to set the file to.
' @remarks All attributes are set at once. In order to modify a single attribute,
' the remaining attribute settings need to be included with the modified attribute.
'
Public Property Let Attributes(ByVal RHS As FileAttributes)
    If SetFileAttributes(mFileName, RHS) = BOOL_FALSE Then IOError Err.LastDllError, mFileName
    MarkDirty
End Property

''
' Returns the creation time as a cDateTime object.
'
' @return The creation time of the file.
'
Public Property Get CreationTime() As Variant
    If mIsDirty Then Refresh
    Set CreationTime = mCreationTime
End Property

''
' Sets the creation time of the file.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let CreationTime(ByVal RHS As Variant)
    Set Me.CreationTime = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the creation time of the file.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set CreationTime(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "CreationTime")

    Set mCreationTime = RHS
    file.SetCreationTime mFileName, mCreationTime
    MarkDirty
End Property

''
' Returns the creation time as a UTC time.
'
' @return The creation time of the file in UTC time.
'
Public Property Get CreationTimeUtc() As Variant
    If mIsDirty Then Refresh
    Set CreationTimeUtc = mCreationTime.ToUniversalTime
End Property

''
' Sets the creation time of the file in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let CreationTimeUtc(ByVal RHS As Variant)
    Set Me.CreationTimeUtc = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the creation time of the file in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set CreationTimeUtc(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "CreationTimeUtc")

    file.SetCreationTimeUtc mFileName, RHS
    MarkDirty
End Property

''
' Returns the last time the file was accessed.
'
' @return A cDateTime object of the last access time.
'
Public Property Get LastAccessTime() As Variant
    If mIsDirty Then Refresh
    Set LastAccessTime = mLastAccessTime
End Property

''
' Sets the time the file was last accessed.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastAccessTime(ByVal RHS As Variant)
    Set Me.LastAccessTime = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the time the file was last accessed.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastAccessTime(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTime")
        
    Set mLastAccessTime = RHS
    file.SetLastAccessTime mFileName, mLastAccessTime
    MarkDirty
End Property

''
' Returns the time the file was last accessed in UTC time.
'
' @return A cDateTime object of the last accessed UTC time.
'
Public Property Get LastAccessTimeUtc() As Variant
    If mIsDirty Then Refresh
    Set LastAccessTimeUtc = mLastAccessTime.ToUniversalTime
End Property

''
' Sets the time the file was last accessed in UTC time.
'
' @param RHS A Date or cDateTime object of the new UTC time.
'
Public Property Let LastAccessTimeUtc(ByVal RHS As Variant)
    Set Me.LastAccessTimeUtc = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the time the file was last accessed in UTC time.
'
' @param RHS A Date or cDateTime object of the new UTC time.

Public Property Set LastAccessTimeUtc(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTimeUtc")
    
    file.SetLastAccessTimeUtc mFileName, RHS
    MarkDirty
End Property

''
' Returns the last time the file was written to.
'
' @return A cDateTime object of the last time the file was written to.
'
Public Property Get LastWriteTime() As Variant
    If mIsDirty Then Refresh
    Set LastWriteTime = mLastWriteTime
End Property

''
' Sets the last time the file was written to.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastWriteTime(ByVal RHS As Variant)
    Set Me.LastAccessTime = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the last time the file was written to.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastWriteTime(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastWriteTime")
    
    Set mLastWriteTime = RHS
    file.SetLastWriteTime mFileName, mLastWriteTime
    MarkDirty
End Property

''
' Returns the last time the file was accessed in UTC time.
'
' @return A cDateTime object of the last time the file was written to.
'
Public Property Get LastWriteTimeUtc() As Variant
    If mIsDirty Then Refresh
    Set LastWriteTimeUtc = mLastWriteTime.ToUniversalTime
End Property

''
' Sets the last the time file was written to in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastWriteTimeUtc(ByVal RHS As Variant)
    Set Me.LastWriteTimeUtc = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the last time the file was written to in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastWriteTimeUtc(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTimeUtc")
    
    file.SetLastWriteTimeUtc mFileName, RHS
    MarkDirty
End Property

''
' Returns a DirectoryInfo object of the directory portion of the file.
'
' @return A DirectoryInfo object of the directory.
' @remarks This does not verify the file exists, or the directory.
'
Public Property Get Directory() As DirectoryInfo
    Set Directory = Cor.NewDirectoryInfo(DirectoryName)
End Property

''
' Returns the directory portion of the filename.
'
' @return The directory portion of the filename.
' @remarks This does not very the existence of the file.
'
Public Property Get DirectoryName() As String
    DirectoryName = Path.GetDirectoryName(mFileName)
End Property

''
' Indicates if the file exists on disk.
'
' @return An indication of the existence of the file.
'
Public Property Get Exists() As Boolean
    Exists = file.Exists(mFileName)
End Property

''
' Returns the extension portion of the filename.
'
' @return The extension portion of the filename.
'
Public Property Get Extension() As String
    Extension = Path.GetExtension(mFileName)
End Property

''
' Returns a full path to the file, including the filename.
'
' @return A full path to the file.
' @remarks If not root directory portion is found in the
' filename, then the currenct directory is used.
'
Public Property Get FullName() As String
    FullName = Path.GetFullPath(mFileName)
End Property

''
' Returns the size of the file in bytes.
'
' @return The size of the file.
' @remarks A Currency datatype is used to allow file sizes greater than 2gig.
'
Public Property Get Length() As Currency
    If mIsDirty Then Refresh
    Length = mFileSize
End Property

''
' Returns the filename portion of the path.
'
' @return The filename and extension.
'
Public Property Get Name() As String
    Name = Path.GetFileName(mFileName)
End Property

''
' Reloads the file attributes and times.
'
' @remarks This will verify that the file exists.
'
Public Sub Refresh()
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    Dim e As Long
    Dim ptr As Long
    
    e = file.GetFileData(mFileName, Data)
    If e <> NO_ERROR Then IOError e, mFileName
    
    With Data
        mFileAttributes = .dwFileAttributes
        Set mCreationTime = cDateTime.FromFileTime(.ftCreationTime)
        Set mLastAccessTime = cDateTime.FromFileTime(.ftLastAccessTime)
        Set mLastWriteTime = cDateTime.FromFileTime(.ftLastWriteTime)
    End With
    
    ' The nFileSizeLow and hFileSizeHigh are in reverse order in
    ' WIN32_FIND_DATA, so we need to manually reverse the 4 byte sections.
    ptr = VarPtr(mFileSize)
    MemLong(ptr) = Data.nFileSizeLow
    MemLong(ptr + 4) = Data.nFileSizeHigh
    mFileSize = mFileSize * 10000@
    
    mIsDirty = False
End Sub

''
' Opens a StreamWriter for the file.
'
' @return A Stream writer to the file.
'
Public Function AppendText() As StreamWriter
    Set AppendText = file.AppendText(mFileName)
End Function

''
' Copies this file to a new location, with optional overwrite.
'
' @param DestFileName The new filename this file will be copied as.
' @param OverWrite Permission to overwrite an existing file.
' @remarks The destination must be on the same volume as this file.
'
Public Sub CopyTo(ByVal DestFileName As String, Optional ByVal OverWrite As Boolean)
    file.Copy mFileName, DestFileName, OverWrite
End Sub

''
' Creates this file and returns a FileStream object to access it.
'
' @return A FileStream object used to access the created file.
' @remarks If the file does not exist it will be created. If it already exists, it will be overwritten.
'
Public Function Create() As FileStream
    Set Create = file.Create(mFileName)
End Function

''
' Returns a StreamWriter to write to this file.
'
' @return The StreamWriter for writing to this file.
' @remarks If the file does not exist it will be created. If it already exists, it will be overwritten.
'
Public Function CreateText() As StreamWriter
    Set CreateText = file.CreateText(mFileName)
End Function

''
' Deletes this file.
'
Public Sub Delete()
    file.Delete mFileName
End Sub

''
' Moves this file to a new location.
'
' @param DestFileName The new location and name of the file.
' @remarks The new location must be on the same volume as the source filename.
'
Public Sub MoveTo(ByVal DestFileName As String)
    file.Move mFileName, DestFileName
End Sub

''
' Opens the file as a FileStream.
'
' @param Mode The method of opening the file.
' @param Access The read/write operations allowed by this FileStream.
' @param Share The read/write operations allows by other processes.
' @return A FileStream object to this file.
'
Public Function OpenFile(ByVal Mode As FileMode, Optional ByVal Access As FileAccess = ReadWriteAccess, Optional ByVal Share As FileShare = FileShare.None) As FileStream
    Set OpenFile = file.OpenFile(mFileName, Mode, Access, Share)
End Function

''
' Returns a FileStream in ReadOnly mode.
'
' @return A FileStream that has ReadOnly access to the file.
'
Public Function OpenRead() As FileStream
    Set OpenRead = file.OpenRead(mFileName)
End Function

''
' Returns a StreamReader to this file.
'
' @return A StreamReader for reading the file.
' @remarks The StreamReader's default encoding is used.
'
Public Function OpenText() As StreamReader
    Set OpenText = file.OpenText(mFileName)
End Function

''
' Returns a FileStream with write access to the file.
'
' @return A FileStream used to write to the file.
'
Public Function OpenWrite() As FileStream
    Set OpenWrite = file.OpenWrite(mFileName)
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = mFileName
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal FileName As String)
    mOriginalFileName = FileName
    mFileName = Path.GetFullPath(FileName)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MarkDirty()
    mIsDirty = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    MarkDirty
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    cObject_Equals = Equals(Value)
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   FileSystemInfo Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get FileSystemInfo_Attributes() As FileAttributes
    FileSystemInfo_Attributes = Attributes
End Property

Private Property Let FileSystemInfo_Attriubtes(ByVal RHS As FileAttributes)
    Attributes = RHS
End Property

Private Property Set FileSystemInfo_CreationTime(ByVal RHS As Variant)
    Set CreationTime = RHS
End Property

Private Property Let FileSystemInfo_CreationTime(ByVal RHS As Variant)
    CreationTime = RHS
End Property

Private Property Get FileSystemInfo_CreationTime() As Variant
    Set FileSystemInfo_CreationTime = CreationTime
End Property

Private Property Set FileSystemInfo_CreationTimeUtc(ByVal RHS As Variant)
    Set CreationTimeUtc = RHS
End Property

Private Property Let FileSystemInfo_CreationTimeUtc(ByVal RHS As Variant)
    CreationTimeUtc = RHS
End Property

Private Property Get FileSystemInfo_CreationTimeUtc() As Variant
    Set FileSystemInfo_CreationTimeUtc = CreationTimeUtc
End Property

Private Sub FileSystemInfo_Delete()
    Delete
End Sub

Private Function FileSystemInfo_Equals(Value As Variant) As Boolean
    FileSystemInfo_Equals = Equals(Value)
End Function

Private Property Get FileSystemInfo_Exists() As Boolean
    FileSystemInfo_Exists = Exists
End Property

Private Property Get FileSystemInfo_Extension() As String
    FileSystemInfo_Extension = Extension
End Property

Private Property Get FileSystemInfo_FullName() As String
    FileSystemInfo_FullName = FullName
End Property

Private Function FileSystemInfo_GetHashCode() As Long
    FileSystemInfo_GetHashCode = GetHashCode
End Function

Private Property Set FileSystemInfo_LastAccessTime(ByVal RHS As Variant)
    Set LastAccessTime = RHS
End Property

Private Property Get FileSystemInfo_LastAccessTime() As Variant
    Set FileSystemInfo_LastAccessTime = LastAccessTime
End Property

Private Property Let FileSystemInfo_LastAccessTime(ByVal RHS As Variant)
    LastAccessTime = RHS
End Property

Private Property Set FileSystemInfo_LastAccessTimeUtc(ByVal RHS As Variant)
    Set LastAccessTimeUtc = RHS
End Property

Private Property Let FileSystemInfo_LastAccessTimeUtc(ByVal RHS As Variant)
    LastAccessTimeUtc = RHS
End Property

Private Property Get FileSystemInfo_LastAccessTimeUtc() As Variant
    Set FileSystemInfo_LastAccessTimeUtc = LastAccessTimeUtc
End Property

Private Property Set FileSystemInfo_LastWriteTime(ByVal RHS As Variant)
    Set LastWriteTime = RHS
End Property

Private Property Let FileSystemInfo_LastWriteTime(ByVal RHS As Variant)
    LastWriteTime = RHS
End Property

Private Property Get FileSystemInfo_LastWriteTime() As Variant
    Set FileSystemInfo_LastWriteTime = LastWriteTime
End Property

Private Property Set FileSystemInfo_LastWriteTimeUtc(ByVal RHS As Variant)
    Set LastWriteTimeUtc = RHS
End Property

Private Property Let FileSystemInfo_LastWriteTimeUtc(ByVal RHS As Variant)
    LastWriteTimeUtc = RHS
End Property

Private Property Get FileSystemInfo_LastWriteTimeUtc() As Variant
    Set FileSystemInfo_LastWriteTimeUtc = LastWriteTimeUtc
End Property

Private Property Get FileSystemInfo_Name() As String
    FileSystemInfo_Name = Name
End Property

Private Sub FileSystemInfo_Refresh()
    Refresh
End Sub

Private Function FileSystemInfo_ToString() As String
    FileSystemInfo_ToString = ToString
End Function
