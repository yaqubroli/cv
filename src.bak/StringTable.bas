Attribute VB_Name = "StringTable"
Option Explicit

Public Const ID_SELF As String = "___self___"
Public Const ID_TYPE As String = "___type___"

Public Const MESSAGE_ERROR_GENERIC As String = "YAML Error"
Public Const MESSAGE_MALFORMED_TYPE As String = "Malformed YAML code on line "
Public Const MESSAGE_MALFORMED_YAML As String = "Malformed type error - this is a problem with the internal dictionary"
Public Const MESSAGE_GETPROP_NOT_STR As String = "Your module has tried to use getProp(), which is meant for type String, on a "
Public Const MESSAGE_GETPROP_NOT_FOUND As String = "Property not found."
