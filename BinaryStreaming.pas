{-------------------------------------------------------------------------------

  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.

-------------------------------------------------------------------------------}
{===============================================================================

  Binary streaming

    Set of functions and classes designed to help writing and reading of binary
    data into/from streams and memory locations.

    All primitives larger than one byte (integers, floats, wide characters -
    including in UTF-16 strings, ...) are always written with little endiannes.

    Boolean data are written as one byte, where zero represents false and any
    non-zero value represents true.

    All strings are written with explicit length, and without terminating zero
    character. First a length is stored, followed immediately by the character
    stream.

    Short strings are stored with 8bit unsigned integer length, all other
    strings are stored with 32bit signed integer length (note that the length
    of 0 or lower marks an empty string, such length is not followed by any
    character belonging to that particular string). Length is in characters,
    not bytes or code points.

    Default type Char is always stored as 16bit unsigned integer, irrespective
    of how it is declared.

    Default type String is always stored as UTF-8 encoded string, again
    irrespective of how it is declared.

    Buffers and array of bytes are both stored as plain byte streams, without
    explicit size.

    Variants are stored in a litle more complex way - first a byte denoting
    the type of the variant is stored (note that value of this byte do NOT
    correspond to TVarType value), directly followed by the value itself.
    The value is stored as if it was a normal variable (eg. for varWord variant
    the function [Ptr/Stream]_WriteUInt16 is called).

      NOTE - given the implementation, variant being written into memory
             (function Ptr_WriteVariant) must not be larger than 2GiB-1 bytes.

    For variant arrays, immediately after the type byte a dimension count is
    stored (int32), followed by an array of indices bounds (two int32 values
    for each dimension, first low bound followed by a high bound, ordered from
    lowest dimension to highest). After this an array of items is stored, each
    item stored as a separate variant. When saving the items, the right-most
    index is incremented first.

      NOTE - only standard variant types are supported, and from them only the
             "streamable" types (boolean, integers, floats, strings) and their
             arrays are allowed - so no empty/null vars, interfaces, objects,
             errors, and so on.

    Also, since older compilers do not support all in-here used variant types
    (namely varUInt64 and varUString), there are some specific rules in effect
    when saving and loading such types in programs compiled by those
    compilers...

      Unsupported types are not saved at all, simply because a variant of that
      type cannot be even created (note that the type UInt64 is sometimes
      declared only an alias for Int64, then it will be saved as Int64).

      When loading unsupported varUInt64, it will be loaded as Int64 (with the
      same BINARY data, ie. not necessarily the same numerical value).

      When loading unsupported varUString (unicode string), it is normally
      loaded, but the returned variant will be of type varOleStr (wide string).

    Parameter Advance in writing and reading functions indicates whether the
    position in stream being written to or read from (or the passed memory
    pointer) can be advanced by number of bytes written or read. When set to
    true, the position (pointer) is advanced, when false, the position is the
    same after the call as was before it.

    Return value of almost all reading and writing functions is number of bytes
    written or read. The exception to this are read functions that are directly
    returning the value being read.

  Version 1.9 (2023-01-22)

  Last change 2023-01-22

  ©2015-2023 František Milt

  Contacts:
    František Milt: frantisek.milt@gmail.com

  Support:
    If you find this code useful, please consider supporting its author(s) by
    making a small donation using the following link(s):

      https://www.paypal.me/FMilt

  Changelog:
    For detailed changelog and history please refer to this git repository:

      github.com/TheLazyTomcat/Lib.BinaryStreaming

  Dependencies:
    AuxClasses         - github.com/TheLazyTomcat/Lib.AuxClasses  
    AuxTypes           - github.com/TheLazyTomcat/Lib.AuxTypes
    StaticMemoryStream - github.com/TheLazyTomcat/Lib.StaticMemoryStream
    StrRect            - github.com/TheLazyTomcat/Lib.StrRect

===============================================================================}
unit BinaryStreaming;

{$IFDEF FPC}
  {$MODE ObjFPC}
  {$INLINE ON}
  {$DEFINE CanInline}
  {$DEFINE CanInlineFPC}
  {$DEFINE FPC_DisableWarns}
  {$MACRO ON}
{$ELSE}
  {$IF CompilerVersion >= 17 then}  // Delphi 2005+
    {$DEFINE CanInline}
  {$IFEND}
{$ENDIF}
{$H+}

interface

uses
  SysUtils, Classes,
  AuxTypes, AuxClasses{$IFNDEF FPC}, StrRect{$ENDIF};

{===============================================================================
    Library-specific exceptions
===============================================================================}
type
  EBSException = class(Exception);

  EBSIndexOutOfBounds   = class(EBSException);
  EBSUnsupportedVarType = class(EBSException);

{===============================================================================
    Endianess - declaration
===============================================================================}
type
  TEndian = (endLittle,endBig,endSystem,endDefault{endLittle});

const
  SysEndian = {$IFDEF ENDIAN_BIG}endBig{$ELSE}endLittle{$ENDIF};

{
  ResolveEndian

  Resolves enSystem to a value of constant SysEndian and enDefault to enLittle,
  other values are returned unchanged.
}
Function ResolveEndian(Endian: TEndian): TEndian;


{===============================================================================
--------------------------------------------------------------------------------
                               Allocation helpers
--------------------------------------------------------------------------------
===============================================================================}
{
  Following functions are returning number of bytes that are required to store
  a given object. They are here mainly to ease allocation when streaming into
  memory.
  Basic types have static size (for completeness sake they are nevertheless
  included), but strings and variants are a little trickier and must be scanned
  to obtain the true size.
}

Function StreamedSize_Bool: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Boolean: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function StreamedSize_Int8: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UInt8: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Int16: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UInt16: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Int32: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UInt32: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Int64: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UInt64: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function StreamedSize_Float32: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Float64: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Float80: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_DateTime: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Currency: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function StreamedSize_AnsiChar: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UTF8Char: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_WideChar: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UnicodeChar: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UCS4Char: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Char: TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function StreamedSize_ShortString(const Str: ShortString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_AnsiString(const Str: AnsiString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UTF8String(const Str: UTF8String): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_WideString(const Str: WideString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UnicodeString(const Str: UnicodeString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UCS4String(const Str: UCS4String): TMemSize;
Function StreamedSize_String(const Str: String): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function StreamedSize_Buffer(Size: TMemSize): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Bytes(Count: TMemSize): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function StreamedSize_Variant(const Value: Variant): TMemSize;


{===============================================================================
--------------------------------------------------------------------------------
                                 Memory writing
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function Ptr_WriteBool_LE(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteBool_LE(Dest: Pointer; Value: ByteBool): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteBool_BE(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteBool_BE(Dest: Pointer; Value: ByteBool): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteBool(var Dest: Pointer; Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteBool(Dest: Pointer; Value: ByteBool; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteBoolean_LE(var Dest: Pointer; Value: Boolean; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteBoolean_LE(Dest: Pointer; Value: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteBoolean_BE(var Dest: Pointer; Value: Boolean; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteBoolean_BE(Dest: Pointer; Value: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteBoolean(var Dest: Pointer; Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteBoolean(Dest: Pointer; Value: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function Ptr_WriteInt8_LE(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteInt8_LE(Dest: Pointer; Value: Int8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteInt8_BE(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteInt8_BE(Dest: Pointer; Value: Int8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteInt8(var Dest: Pointer; Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteInt8(Dest: Pointer; Value: Int8; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt8_LE(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteUInt8_LE(Dest: Pointer; Value: UInt8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteUInt8_BE(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteUInt8_BE(Dest: Pointer; Value: UInt8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteUInt8(var Dest: Pointer; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUInt8(Dest: Pointer; Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteInt16_LE(var Dest: Pointer; Value: Int16; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt16_LE(Dest: Pointer; Value: Int16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt16_BE(var Dest: Pointer; Value: Int16; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt16_BE(Dest: Pointer; Value: Int16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt16(var Dest: Pointer; Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteInt16(Dest: Pointer; Value: Int16; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt16_LE(var Dest: Pointer; Value: UInt16; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt16_LE(Dest: Pointer; Value: UInt16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt16_BE(var Dest: Pointer; Value: UInt16; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt16_BE(Dest: Pointer; Value: UInt16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt16(var Dest: Pointer; Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUInt16(Dest: Pointer; Value: UInt16; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteInt32_LE(var Dest: Pointer; Value: Int32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt32_LE(Dest: Pointer; Value: Int32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt32_BE(var Dest: Pointer; Value: Int32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt32_BE(Dest: Pointer; Value: Int32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt32(var Dest: Pointer; Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteInt32(Dest: Pointer; Value: Int32; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt32_LE(var Dest: Pointer; Value: UInt32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt32_LE(Dest: Pointer; Value: UInt32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt32_BE(var Dest: Pointer; Value: UInt32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt32_BE(Dest: Pointer; Value: UInt32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt32(var Dest: Pointer; Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUInt32(Dest: Pointer; Value: UInt32; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteInt64_LE(var Dest: Pointer; Value: Int64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt64_LE(Dest: Pointer; Value: Int64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt64_BE(var Dest: Pointer; Value: Int64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt64_BE(Dest: Pointer; Value: Int64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt64(var Dest: Pointer; Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteInt64(Dest: Pointer; Value: Int64; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt64_LE(var Dest: Pointer; Value: UInt64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt64_LE(Dest: Pointer; Value: UInt64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt64_BE(var Dest: Pointer; Value: UInt64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt64_BE(Dest: Pointer; Value: UInt64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt64(var Dest: Pointer; Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUInt64(Dest: Pointer; Value: UInt64; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Ptr_WriteFloat32_LE(var Dest: Pointer; Value: Float32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat32_LE(Dest: Pointer; Value: Float32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat32_BE(var Dest: Pointer; Value: Float32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat32_BE(Dest: Pointer; Value: Float32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat32(var Dest: Pointer; Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteFloat32(Dest: Pointer; Value: Float32; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteFloat64_LE(var Dest: Pointer; Value: Float64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat64_LE(Dest: Pointer; Value: Float64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat64_BE(var Dest: Pointer; Value: Float64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat64_BE(Dest: Pointer; Value: Float64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat64(var Dest: Pointer; Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteFloat64(Dest: Pointer; Value: Float64; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteFloat80_LE(var Dest: Pointer; Value: Float80; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat80_LE(Dest: Pointer; Value: Float80): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat80_BE(var Dest: Pointer; Value: Float80; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat80_BE(Dest: Pointer; Value: Float80): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat80(var Dest: Pointer; Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteFloat80(Dest: Pointer; Value: Float80; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteDateTime_LE(var Dest: Pointer; Value: TDateTime; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteDateTime_LE(Dest: Pointer; Value: TDateTime): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteDateTime_BE(var Dest: Pointer; Value: TDateTime; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteDateTime_BE(Dest: Pointer; Value: TDateTime): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteDateTime(var Dest: Pointer; Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteDateTime(Dest: Pointer; Value: TDateTime; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function Ptr_WriteCurrency_LE(var Dest: Pointer; Value: Currency; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteCurrency_LE(Dest: Pointer; Value: Currency): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteCurrency_BE(var Dest: Pointer; Value: Currency; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteCurrency_BE(Dest: Pointer; Value: Currency): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteCurrency(var Dest: Pointer; Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteCurrency(Dest: Pointer; Value: Currency; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function Ptr_WriteAnsiChar_LE(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteAnsiChar_LE(Dest: Pointer; Value: AnsiChar): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteAnsiChar_BE(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteAnsiChar_BE(Dest: Pointer; Value: AnsiChar): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteAnsiChar(var Dest: Pointer; Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteAnsiChar(Dest: Pointer; Value: AnsiChar; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8Char_LE(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteUTF8Char_LE(Dest: Pointer; Value: UTF8Char): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteUTF8Char_BE(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteUTF8Char_BE(Dest: Pointer; Value: UTF8Char): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteUTF8Char(var Dest: Pointer; Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUTF8Char(Dest: Pointer; Value: UTF8Char; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteWideChar_LE(var Dest: Pointer; Value: WideChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteWideChar_LE(Dest: Pointer; Value: WideChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteWideChar_BE(var Dest: Pointer; Value: WideChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteWideChar_BE(Dest: Pointer; Value: WideChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteWideChar(var Dest: Pointer; Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteWideChar(Dest: Pointer; Value: WideChar; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeChar_LE(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUnicodeChar_LE(Dest: Pointer; Value: UnicodeChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUnicodeChar_BE(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUnicodeChar_BE(Dest: Pointer; Value: UnicodeChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUnicodeChar(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUnicodeChar(Dest: Pointer; Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUCS4Char_LE(var Dest: Pointer; Value: UCS4Char; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUCS4Char_LE(Dest: Pointer; Value: UCS4Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUCS4Char_BE(var Dest: Pointer; Value: UCS4Char; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUCS4Char_BE(Dest: Pointer; Value: UCS4Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUCS4Char(var Dest: Pointer; Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUCS4Char(Dest: Pointer; Value: UCS4Char; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteChar_LE(var Dest: Pointer; Value: Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteChar_LE(Dest: Pointer; Value: Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteChar_BE(var Dest: Pointer; Value: Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteChar_BE(Dest: Pointer; Value: Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteChar(var Dest: Pointer; Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteChar(Dest: Pointer; Value: Char; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function Ptr_WriteShortString_LE(var Dest: Pointer; const Str: ShortString; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteShortString_LE(Dest: Pointer; const Str: ShortString): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteShortString_BE(var Dest: Pointer; const Str: ShortString; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_WriteShortString_BE(Dest: Pointer; const Str: ShortString): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_WriteShortString(var Dest: Pointer; const Str: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteShortString(Dest: Pointer; const Str: ShortString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiString_LE(var Dest: Pointer; const Str: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteAnsiString_LE(Dest: Pointer; const Str: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteAnsiString_BE(var Dest: Pointer; const Str: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteAnsiString_BE(Dest: Pointer; const Str: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteAnsiString(var Dest: Pointer; const Str: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteAnsiString(Dest: Pointer; const Str: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8String_LE(var Dest: Pointer; const Str: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUTF8String_LE(Dest: Pointer; const Str: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUTF8String_BE(var Dest: Pointer; const Str: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUTF8String_BE(Dest: Pointer; const Str: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUTF8String(var Dest: Pointer; const Str: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUTF8String(Dest: Pointer; const Str: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteWideString_LE(var Dest: Pointer; const Str: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteWideString_LE(Dest: Pointer; const Str: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteWideString_BE(var Dest: Pointer; const Str: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteWideString_BE(Dest: Pointer; const Str: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteWideString(var Dest: Pointer; const Str: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteWideString(Dest: Pointer; const Str: WideString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeString_LE(var Dest: Pointer; const Str: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUnicodeString_LE(Dest: Pointer; const Str: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUnicodeString_BE(var Dest: Pointer; const Str: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUnicodeString_BE(Dest: Pointer; const Str: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUnicodeString(var Dest: Pointer; const Str: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUnicodeString(Dest: Pointer; const Str: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteUCS4String_LE(var Dest: Pointer; const Str: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUCS4String_LE(Dest: Pointer; const Str: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUCS4String_BE(var Dest: Pointer; const Str: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUCS4String_BE(Dest: Pointer; const Str: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUCS4String(var Dest: Pointer; const Str: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUCS4String(Dest: Pointer; const Str: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteString_LE(var Dest: Pointer; const Str: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteString_LE(Dest: Pointer; const Str: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteString_BE(var Dest: Pointer; const Str: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteString_BE(Dest: Pointer; const Str: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteString(var Dest: Pointer; const Str: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteString(Dest: Pointer; const Str: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function Ptr_WriteBuffer_LE(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBuffer_LE(Dest: Pointer; const Buffer; Size: TMemSize): TMemSize; overload;

Function Ptr_WriteBuffer_BE(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBuffer_BE(Dest: Pointer; const Buffer; Size: TMemSize): TMemSize; overload;

Function Ptr_WriteBuffer(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteBuffer(Dest: Pointer; const Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------
// note - calls with open array cannot be inlined both in Delphi and FPC

Function Ptr_WriteBytes_LE(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBytes_LE(Dest: Pointer; const Value: array of UInt8): TMemSize; overload;

Function Ptr_WriteBytes_BE(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBytes_BE(Dest: Pointer; const Value: array of UInt8): TMemSize; overload;

Function Ptr_WriteBytes(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteBytes(Dest: Pointer; const Value: array of UInt8; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_FillBytes_LE(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_FillBytes_LE(Dest: Pointer; Count: TMemSize; Value: UInt8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_FillBytes_BE(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_FillBytes_BE(Dest: Pointer; Count: TMemSize; Value: UInt8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_FillBytes(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_FillBytes(Dest: Pointer; Count: TMemSize; Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Ptr_WriteVariant_LE(var Dest: Pointer; const Value: Variant; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteVariant_LE(Dest: Pointer; const Value: Variant): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteVariant_BE(var Dest: Pointer; const Value: Variant; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteVariant_BE(Dest: Pointer; const Value: Variant): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteVariant(var Dest: Pointer; const Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteVariant(Dest: Pointer; const Value: Variant; Endian: TEndian = endDefault): TMemSize; overload;


{===============================================================================
--------------------------------------------------------------------------------
                                 Memory reading
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function Ptr_ReadBool_LE(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadBool_LE(Src: Pointer; out Value: ByteBool): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadBool_BE(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadBool_BE(Src: Pointer; out Value: ByteBool): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadBool(var Src: Pointer; out Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadBool(Src: Pointer; out Value: ByteBool; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetBool_LE(var Src: Pointer; Advance: Boolean): ByteBool; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetBool_LE(Src: Pointer): ByteBool; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetBool_BE(var Src: Pointer; Advance: Boolean): ByteBool; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetBool_BE(Src: Pointer): ByteBool; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetBool(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): ByteBool; overload;
Function Ptr_GetBool(Src: Pointer; Endian: TEndian = endDefault): ByteBool; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadBoolean_LE(var Src: Pointer; out Value: Boolean; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadBoolean_LE(Src: Pointer; out Value: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadBoolean_BE(var Src: Pointer; out Value: Boolean; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadBoolean_BE(Src: Pointer; out Value: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadBoolean(var Src: Pointer; out Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadBoolean(Src: Pointer; out Value: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetBoolean_LE(var Src: Pointer; Advance: Boolean): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetBoolean_LE(Src: Pointer): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetBoolean_BE(var Src: Pointer; Advance: Boolean): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetBoolean_BE(Src: Pointer): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetBoolean(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetBoolean(Src: Pointer; Endian: TEndian = endDefault): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function Ptr_ReadInt8_LE(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadInt8_LE(Src: Pointer; out Value: Int8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadInt8_BE(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadInt8_BE(Src: Pointer; out Value: Int8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadInt8(var Src: Pointer; out Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadInt8(Src: Pointer; out Value: Int8; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetInt8_LE(var Src: Pointer; Advance: Boolean): Int8; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetInt8_LE(Src: Pointer): Int8; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetInt8_BE(var Src: Pointer; Advance: Boolean): Int8; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetInt8_BE(Src: Pointer): Int8; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetInt8(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int8; overload;
Function Ptr_GetInt8(Src: Pointer; Endian: TEndian = endDefault): Int8; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt8_LE(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadUInt8_LE(Src: Pointer; out Value: UInt8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadUInt8_BE(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadUInt8_BE(Src: Pointer; out Value: UInt8): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadUInt8(var Src: Pointer; out Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUInt8(Src: Pointer; out Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUInt8_LE(var Src: Pointer; Advance: Boolean): UInt8; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetUInt8_LE(Src: Pointer): UInt8; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetUInt8_BE(var Src: Pointer; Advance: Boolean): UInt8; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetUInt8_BE(Src: Pointer): UInt8; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetUInt8(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt8; overload;
Function Ptr_GetUInt8(Src: Pointer; Endian: TEndian = endDefault): UInt8; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadInt16_LE(var Src: Pointer; out Value: Int16; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt16_LE(Src: Pointer; out Value: Int16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt16_BE(var Src: Pointer; out Value: Int16; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt16_BE(Src: Pointer; out Value: Int16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt16(var Src: Pointer; out Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadInt16(Src: Pointer; out Value: Int16; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetInt16_LE(var Src: Pointer; Advance: Boolean): Int16; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt16_LE(Src: Pointer): Int16; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt16_BE(var Src: Pointer; Advance: Boolean): Int16; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt16_BE(Src: Pointer): Int16; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt16(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int16; overload;
Function Ptr_GetInt16(Src: Pointer; Endian: TEndian = endDefault): Int16; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt16_LE(var Src: Pointer; out Value: UInt16; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt16_LE(Src: Pointer; out Value: UInt16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt16_BE(var Src: Pointer; out Value: UInt16; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt16_BE(Src: Pointer; out Value: UInt16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt16(var Src: Pointer; out Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUInt16(Src: Pointer; out Value: UInt16; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUInt16_LE(var Src: Pointer; Advance: Boolean): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt16_LE(Src: Pointer): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt16_BE(var Src: Pointer; Advance: Boolean): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt16_BE(Src: Pointer): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt16(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt16; overload;
Function Ptr_GetUInt16(Src: Pointer; Endian: TEndian = endDefault): UInt16; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadInt32_LE(var Src: Pointer; out Value: Int32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt32_LE(Src: Pointer; out Value: Int32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt32_BE(var Src: Pointer; out Value: Int32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt32_BE(Src: Pointer; out Value: Int32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt32(var Src: Pointer; out Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadInt32(Src: Pointer; out Value: Int32; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetInt32_LE(var Src: Pointer; Advance: Boolean): Int32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt32_LE(Src: Pointer): Int32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt32_BE(var Src: Pointer; Advance: Boolean): Int32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt32_BE(Src: Pointer): Int32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int32; overload;
Function Ptr_GetInt32(Src: Pointer; Endian: TEndian = endDefault): Int32; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt32_LE(var Src: Pointer; out Value: UInt32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt32_LE(Src: Pointer; out Value: UInt32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt32_BE(var Src: Pointer; out Value: UInt32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt32_BE(Src: Pointer; out Value: UInt32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt32(var Src: Pointer; out Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUInt32(Src: Pointer; out Value: UInt32; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUInt32_LE(var Src: Pointer; Advance: Boolean): UInt32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt32_LE(Src: Pointer): UInt32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt32_BE(var Src: Pointer; Advance: Boolean): UInt32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt32_BE(Src: Pointer): UInt32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt32; overload;
Function Ptr_GetUInt32(Src: Pointer; Endian: TEndian = endDefault): UInt32; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadInt64_LE(var Src: Pointer; out Value: Int64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt64_LE(Src: Pointer; out Value: Int64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt64_BE(var Src: Pointer; out Value: Int64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt64_BE(Src: Pointer; out Value: Int64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt64(var Src: Pointer; out Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadInt64(Src: Pointer; out Value: Int64; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetInt64_LE(var Src: Pointer; Advance: Boolean): Int64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt64_LE(Src: Pointer): Int64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt64_BE(var Src: Pointer; Advance: Boolean): Int64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt64_BE(Src: Pointer): Int64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int64; overload;
Function Ptr_GetInt64(Src: Pointer; Endian: TEndian = endDefault): Int64; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt64_LE(var Src: Pointer; out Value: UInt64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt64_LE(Src: Pointer; out Value: UInt64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt64_BE(var Src: Pointer; out Value: UInt64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt64_BE(Src: Pointer; out Value: UInt64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt64(var Src: Pointer; out Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUInt64(Src: Pointer; out Value: UInt64; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUInt64_LE(var Src: Pointer; Advance: Boolean): UInt64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt64_LE(Src: Pointer): UInt64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt64_BE(var Src: Pointer; Advance: Boolean): UInt64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt64_BE(Src: Pointer): UInt64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt64; overload;
Function Ptr_GetUInt64(Src: Pointer; Endian: TEndian = endDefault): UInt64; overload;

{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Ptr_ReadFloat32_LE(var Src: Pointer; out Value: Float32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat32_LE(Src: Pointer; out Value: Float32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat32_BE(var Src: Pointer; out Value: Float32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat32_BE(Src: Pointer; out Value: Float32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat32(var Src: Pointer; out Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadFloat32(Src: Pointer; out Value: Float32; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetFloat32_LE(var Src: Pointer; Advance: Boolean): Float32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat32_LE(Src: Pointer): Float32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat32_BE(var Src: Pointer; Advance: Boolean): Float32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat32_BE(Src: Pointer): Float32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float32; overload;
Function Ptr_GetFloat32(Src: Pointer; Endian: TEndian = endDefault): Float32; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadFloat64_LE(var Src: Pointer; out Value: Float64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat64_LE(Src: Pointer; out Value: Float64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat64_BE(var Src: Pointer; out Value: Float64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat64_BE(Src: Pointer; out Value: Float64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat64(var Src: Pointer; out Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadFloat64(Src: Pointer; out Value: Float64; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetFloat64_LE(var Src: Pointer; Advance: Boolean): Float64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat64_LE(Src: Pointer): Float64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat64_BE(var Src: Pointer; Advance: Boolean): Float64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat64_BE(Src: Pointer): Float64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float64; overload;
Function Ptr_GetFloat64(Src: Pointer; Endian: TEndian = endDefault): Float64; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadFloat80_LE(var Src: Pointer; out Value: Float80; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat80_LE(Src: Pointer; out Value: Float80): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat80_BE(var Src: Pointer; out Value: Float80; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat80_BE(Src: Pointer; out Value: Float80): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat80(var Src: Pointer; out Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadFloat80(Src: Pointer; out Value: Float80; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetFloat80_LE(var Src: Pointer; Advance: Boolean): Float80; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat80_LE(Src: Pointer): Float80; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat80_BE(var Src: Pointer; Advance: Boolean): Float80; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat80_BE(Src: Pointer): Float80; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat80(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float80; overload;
Function Ptr_GetFloat80(Src: Pointer; Endian: TEndian = endDefault): Float80; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadDateTime_LE(var Src: Pointer; out Value: TDateTime; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadDateTime_LE(Src: Pointer; out Value: TDateTime): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadDateTime_BE(var Src: Pointer; out Value: TDateTime; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadDateTime_BE(Src: Pointer; out Value: TDateTime): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadDateTime(var Src: Pointer; out Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadDateTime(Src: Pointer; out Value: TDateTime; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetDateTime_LE(var Src: Pointer; Advance: Boolean): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetDateTime_LE(Src: Pointer): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetDateTime_BE(var Src: Pointer; Advance: Boolean): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetDateTime_BE(Src: Pointer): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetDateTime(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetDateTime(Src: Pointer; Endian: TEndian = endDefault): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function Ptr_ReadCurrency_LE(var Src: Pointer; out Value: Currency; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadCurrency_LE(Src: Pointer; out Value: Currency): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadCurrency_BE(var Src: Pointer; out Value: Currency; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadCurrency_BE(Src: Pointer; out Value: Currency): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadCurrency(var Src: Pointer; out Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadCurrency(Src: Pointer; out Value: Currency; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetCurrency_LE(var Src: Pointer; Advance: Boolean): Currency; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetCurrency_LE(Src: Pointer): Currency; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetCurrency_BE(var Src: Pointer; Advance: Boolean): Currency; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetCurrency_BE(Src: Pointer): Currency; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetCurrency(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Currency; overload;
Function Ptr_GetCurrency(Src: Pointer; Endian: TEndian = endDefault): Currency; overload;

{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function Ptr_ReadAnsiChar_LE(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadAnsiChar_LE(Src: Pointer; out Value: AnsiChar): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadAnsiChar_BE(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadAnsiChar_BE(Src: Pointer; out Value: AnsiChar): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadAnsiChar(var Src: Pointer; out Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadAnsiChar(Src: Pointer; out Value: AnsiChar; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetAnsiChar_LE(var Src: Pointer; Advance: Boolean): AnsiChar; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetAnsiChar_LE(Src: Pointer): AnsiChar; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetAnsiChar_BE(var Src: Pointer; Advance: Boolean): AnsiChar; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetAnsiChar_BE(Src: Pointer): AnsiChar; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetAnsiChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): AnsiChar; overload;
Function Ptr_GetAnsiChar(Src: Pointer; Endian: TEndian = endDefault): AnsiChar; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8Char_LE(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadUTF8Char_LE(Src: Pointer; out Value: UTF8Char): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadUTF8Char_BE(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadUTF8Char_BE(Src: Pointer; out Value: UTF8Char): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadUTF8Char(var Src: Pointer; out Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUTF8Char(Src: Pointer; out Value: UTF8Char; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUTF8Char_LE(var Src: Pointer; Advance: Boolean): UTF8Char; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetUTF8Char_LE(Src: Pointer): UTF8Char; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetUTF8Char_BE(var Src: Pointer; Advance: Boolean): UTF8Char; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetUTF8Char_BE(Src: Pointer): UTF8Char; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetUTF8Char(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UTF8Char; overload;
Function Ptr_GetUTF8Char(Src: Pointer; Endian: TEndian = endDefault): UTF8Char; overload;
 
//------------------------------------------------------------------------------

Function Ptr_ReadWideChar_LE(var Src: Pointer; out Value: WideChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadWideChar_LE(Src: Pointer; out Value: WideChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadWideChar_BE(var Src: Pointer; out Value: WideChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadWideChar_BE(Src: Pointer; out Value: WideChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadWideChar(var Src: Pointer; out Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadWideChar(Src: Pointer; out Value: WideChar; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetWideChar_LE(var Src: Pointer; Advance: Boolean): WideChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetWideChar_LE(Src: Pointer): WideChar; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetWideChar_BE(var Src: Pointer; Advance: Boolean): WideChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetWideChar_BE(Src: Pointer): WideChar; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetWideChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): WideChar; overload;
Function Ptr_GetWideChar(Src: Pointer; Endian: TEndian = endDefault): WideChar; overload;
 
//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeChar_LE(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUnicodeChar_LE(Src: Pointer; out Value: UnicodeChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUnicodeChar_BE(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUnicodeChar_BE(Src: Pointer; out Value: UnicodeChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUnicodeChar(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUnicodeChar(Src: Pointer; out Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUnicodeChar_LE(var Src: Pointer; Advance: Boolean): UnicodeChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUnicodeChar_LE(Src: Pointer): UnicodeChar; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUnicodeChar_BE(var Src: Pointer; Advance: Boolean): UnicodeChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUnicodeChar_BE(Src: Pointer): UnicodeChar; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUnicodeChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UnicodeChar; overload;
Function Ptr_GetUnicodeChar(Src: Pointer; Endian: TEndian = endDefault): UnicodeChar; overload;
   
//------------------------------------------------------------------------------

Function Ptr_ReadUCS4Char_LE(var Src: Pointer; out Value: UCS4Char; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUCS4Char_LE(Src: Pointer; out Value: UCS4Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUCS4Char_BE(var Src: Pointer; out Value: UCS4Char; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUCS4Char_BE(Src: Pointer; out Value: UCS4Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUCS4Char(var Src: Pointer; out Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUCS4Char(Src: Pointer; out Value: UCS4Char; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUCS4Char_LE(var Src: Pointer; Advance: Boolean): UCS4Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUCS4Char_LE(Src: Pointer): UCS4Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUCS4Char_BE(var Src: Pointer; Advance: Boolean): UCS4Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUCS4Char_BE(Src: Pointer): UCS4Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUCS4Char(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UCS4Char; overload;
Function Ptr_GetUCS4Char(Src: Pointer; Endian: TEndian = endDefault): UCS4Char; overload;
    
//------------------------------------------------------------------------------

Function Ptr_ReadChar_LE(var Src: Pointer; out Value: Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadChar_LE(Src: Pointer; out Value: Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadChar_BE(var Src: Pointer; out Value: Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadChar_BE(Src: Pointer; out Value: Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadChar(var Src: Pointer; out Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadChar(Src: Pointer; out Value: Char; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetChar_LE(var Src: Pointer; Advance: Boolean): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetChar_LE(Src: Pointer): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetChar_BE(var Src: Pointer; Advance: Boolean): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetChar_BE(Src: Pointer): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetChar(Src: Pointer; Endian: TEndian = endDefault): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function Ptr_ReadShortString_LE(var Src: Pointer; out Str: ShortString; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadShortString_LE(Src: Pointer; out Str: ShortString): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadShortString_BE(var Src: Pointer; out Str: ShortString; Advance: Boolean): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_ReadShortString_BE(Src: Pointer; out Str: ShortString): TMemSize; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_ReadShortString(var Src: Pointer; out Str: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadShortString(Src: Pointer; out Str: ShortString; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetShortString_LE(var Src: Pointer; Advance: Boolean): ShortString; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetShortString_LE(Src: Pointer): ShortString; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetShortString_BE(var Src: Pointer; Advance: Boolean): ShortString; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}
Function Ptr_GetShortString_BE(Src: Pointer): ShortString; overload;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Ptr_GetShortString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): ShortString; overload;
Function Ptr_GetShortString(Src: Pointer; Endian: TEndian = endDefault): ShortString; overload;
       
//------------------------------------------------------------------------------

Function Ptr_ReadAnsiString_LE(var Src: Pointer; out Str: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadAnsiString_LE(Src: Pointer; out Str: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadAnsiString_BE(var Src: Pointer; out Str: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadAnsiString_BE(Src: Pointer; out Str: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadAnsiString(var Src: Pointer; out Str: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadAnsiString(Src: Pointer; out Str: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetAnsiString_LE(var Src: Pointer; Advance: Boolean): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetAnsiString_LE(Src: Pointer): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetAnsiString_BE(var Src: Pointer; Advance: Boolean): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetAnsiString_BE(Src: Pointer): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetAnsiString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): AnsiString; overload;
Function Ptr_GetAnsiString(Src: Pointer; Endian: TEndian = endDefault): AnsiString; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8String_LE(var Src: Pointer; out Str: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUTF8String_LE(Src: Pointer; out Str: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUTF8String_BE(var Src: Pointer; out Str: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUTF8String_BE(Src: Pointer; out Str: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUTF8String(var Src: Pointer; out Str: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUTF8String(Src: Pointer; out Str: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUTF8String_LE(var Src: Pointer; Advance: Boolean): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUTF8String_LE(Src: Pointer): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUTF8String_BE(var Src: Pointer; Advance: Boolean): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUTF8String_BE(Src: Pointer): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUTF8String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UTF8String; overload;
Function Ptr_GetUTF8String(Src: Pointer; Endian: TEndian = endDefault): UTF8String; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadWideString_LE(var Src: Pointer; out Str: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadWideString_LE(Src: Pointer; out Str: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadWideString_BE(var Src: Pointer; out Str: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadWideString_BE(Src: Pointer; out Str: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadWideString(var Src: Pointer; out Str: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadWideString(Src: Pointer; out Str: WideString; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetWideString_LE(var Src: Pointer; Advance: Boolean): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetWideString_LE(Src: Pointer): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetWideString_BE(var Src: Pointer; Advance: Boolean): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetWideString_BE(Src: Pointer): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetWideString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): WideString; overload;
Function Ptr_GetWideString(Src: Pointer; Endian: TEndian = endDefault): WideString; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeString_LE(var Src: Pointer; out Str: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUnicodeString_LE(Src: Pointer; out Str: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUnicodeString_BE(var Src: Pointer; out Str: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUnicodeString_BE(Src: Pointer; out Str: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUnicodeString(var Src: Pointer; out Str: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUnicodeString(Src: Pointer; out Str: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUnicodeString_LE(var Src: Pointer; Advance: Boolean): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUnicodeString_LE(Src: Pointer): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUnicodeString_BE(var Src: Pointer; Advance: Boolean): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUnicodeString_BE(Src: Pointer): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUnicodeString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString; overload;
Function Ptr_GetUnicodeString(Src: Pointer; Endian: TEndian = endDefault): UnicodeString; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadUCS4String_LE(var Src: Pointer; out Str: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUCS4String_LE(Src: Pointer; out Str: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUCS4String_BE(var Src: Pointer; out Str: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUCS4String_BE(Src: Pointer; out Str: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUCS4String(var Src: Pointer; out Str: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUCS4String(Src: Pointer; out Str: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetUCS4String_LE(var Src: Pointer; Advance: Boolean): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUCS4String_LE(Src: Pointer): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUCS4String_BE(var Src: Pointer; Advance: Boolean): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUCS4String_BE(Src: Pointer): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUCS4String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UCS4String; overload;
Function Ptr_GetUCS4String(Src: Pointer; Endian: TEndian = endDefault): UCS4String; overload;
      
//------------------------------------------------------------------------------

Function Ptr_ReadString_LE(var Src: Pointer; out Str: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadString_LE(Src: Pointer; out Str: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadString_BE(var Src: Pointer; out Str: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadString_BE(Src: Pointer; out Str: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadString(var Src: Pointer; out Str: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadString(Src: Pointer; out Str: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetString_LE(var Src: Pointer; Advance: Boolean): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetString_LE(Src: Pointer): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetString_BE(var Src: Pointer; Advance: Boolean): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetString_BE(Src: Pointer): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetString(Src: Pointer; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function Ptr_ReadBuffer_LE(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadBuffer_LE(Src: Pointer; out Buffer; Size: TMemSize): TMemSize; overload;

Function Ptr_ReadBuffer_BE(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadBuffer_BE(Src: Pointer; out Buffer; Size: TMemSize): TMemSize; overload;

Function Ptr_ReadBuffer(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadBuffer(Src: Pointer; out Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Ptr_ReadVariant_LE(var Src: Pointer; out Value: Variant; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadVariant_LE(Src: Pointer; out Value: Variant): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadVariant_BE(var Src: Pointer; out Value: Variant; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadVariant_BE(Src: Pointer; out Value: Variant): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadVariant(var Src: Pointer; out Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadVariant(Src: Pointer; out Value: Variant; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetVariant_LE(var Src: Pointer; Advance: Boolean): Variant; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetVariant_LE(Src: Pointer): Variant; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetVariant_BE(var Src: Pointer; Advance: Boolean): Variant; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetVariant_BE(Src: Pointer): Variant; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetVariant(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Variant; overload;
Function Ptr_GetVariant(Src: Pointer; Endian: TEndian = endDefault): Variant; overload;


{===============================================================================
--------------------------------------------------------------------------------
                                 Stream writing
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function Stream_WriteBool_LE(Stream: TStream; Value: ByteBool; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteBool_BE(Stream: TStream; Value: ByteBool; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteBool(Stream: TStream; Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteBool(Stream: TStream; Value: ByteBool; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteBoolean_LE(Stream: TStream; Value: Boolean; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteBoolean_BE(Stream: TStream; Value: Boolean; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteBoolean(Stream: TStream; Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteBoolean(Stream: TStream; Value: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function Stream_WriteInt8_LE(Stream: TStream; Value: Int8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteInt8_BE(Stream: TStream; Value: Int8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteInt8(Stream: TStream; Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteInt8(Stream: TStream; Value: Int8; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUInt8_LE(Stream: TStream; Value: UInt8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteUInt8_BE(Stream: TStream; Value: UInt8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteUInt8(Stream: TStream; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUInt8(Stream: TStream; Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteInt16_LE(Stream: TStream; Value: Int16; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt16_BE(Stream: TStream; Value: Int16; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt16(Stream: TStream; Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteInt16(Stream: TStream; Value: Int16; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUInt16_LE(Stream: TStream; Value: UInt16; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt16_BE(Stream: TStream; Value: UInt16; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt16(Stream: TStream; Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUInt16(Stream: TStream; Value: UInt16; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteInt32_LE(Stream: TStream; Value: Int32; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt32_BE(Stream: TStream; Value: Int32; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt32(Stream: TStream; Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteInt32(Stream: TStream; Value: Int32; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUInt32_LE(Stream: TStream; Value: UInt32; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt32_BE(Stream: TStream; Value: UInt32; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt32(Stream: TStream; Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUInt32(Stream: TStream; Value: UInt32; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteInt64_LE(Stream: TStream; Value: Int64; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt64_BE(Stream: TStream; Value: Int64; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt64(Stream: TStream; Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteInt64(Stream: TStream; Value: Int64; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUInt64_LE(Stream: TStream; Value: UInt64; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt64_BE(Stream: TStream; Value: UInt64; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt64(Stream: TStream; Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUInt64(Stream: TStream; Value: UInt64; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Stream_WriteFloat32_LE(Stream: TStream; Value: Float32; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat32_BE(Stream: TStream; Value: Float32; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat32(Stream: TStream; Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteFloat32(Stream: TStream; Value: Float32; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteFloat64_LE(Stream: TStream; Value: Float64; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat64_BE(Stream: TStream; Value: Float64; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat64(Stream: TStream; Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteFloat64(Stream: TStream; Value: Float64; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteFloat80_LE(Stream: TStream; Value: Float80; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat80_BE(Stream: TStream; Value: Float80; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat80(Stream: TStream; Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteFloat80(Stream: TStream; Value: Float80; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteDateTime_LE(Stream: TStream; Value: TDateTime; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteDateTime_BE(Stream: TStream; Value: TDateTime; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteDateTime(Stream: TStream; Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteDateTime(Stream: TStream; Value: TDateTime; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function Stream_WriteCurrency_LE(Stream: TStream; Value: Currency; Advance: Boolean = True): TMemSize;

Function Stream_WriteCurrency_BE(Stream: TStream; Value: Currency; Advance: Boolean = True): TMemSize;

Function Stream_WriteCurrency(Stream: TStream; Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteCurrency(Stream: TStream; Value: Currency; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function Stream_WriteAnsiChar_LE(Stream: TStream; Value: AnsiChar; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteAnsiChar_BE(Stream: TStream; Value: AnsiChar; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteAnsiChar(Stream: TStream; Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteAnsiChar(Stream: TStream; Value: AnsiChar; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUTF8Char_LE(Stream: TStream; Value: UTF8Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteUTF8Char_BE(Stream: TStream; Value: UTF8Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteUTF8Char(Stream: TStream; Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUTF8Char(Stream: TStream; Value: UTF8Char; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteWideChar_LE(Stream: TStream; Value: WideChar; Advance: Boolean = True): TMemSize;

Function Stream_WriteWideChar_BE(Stream: TStream; Value: WideChar; Advance: Boolean = True): TMemSize;

Function Stream_WriteWideChar(Stream: TStream; Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteWideChar(Stream: TStream; Value: WideChar; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUnicodeChar_LE(Stream: TStream; Value: UnicodeChar; Advance: Boolean = True): TMemSize;

Function Stream_WriteUnicodeChar_BE(Stream: TStream; Value: UnicodeChar; Advance: Boolean = True): TMemSize;

Function Stream_WriteUnicodeChar(Stream: TStream; Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUnicodeChar(Stream: TStream; Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUCS4Char_LE(Stream: TStream; Value: UCS4Char; Advance: Boolean = True): TMemSize;

Function Stream_WriteUCS4Char_BE(Stream: TStream; Value: UCS4Char; Advance: Boolean = True): TMemSize;

Function Stream_WriteUCS4Char(Stream: TStream; Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUCS4Char(Stream: TStream; Value: UCS4Char; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteChar_LE(Stream: TStream; Value: Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteChar_BE(Stream: TStream; Value: Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteChar(Stream: TStream; Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteChar(Stream: TStream; Value: Char; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function Stream_WriteShortString_LE(Stream: TStream; const Str: ShortString; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteShortString_BE(Stream: TStream; const Str: ShortString; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_WriteShortString(Stream: TStream; const Str: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteShortString(Stream: TStream; const Str: ShortString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteAnsiString_LE(Stream: TStream; const Str: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_WriteAnsiString_BE(Stream: TStream; const Str: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_WriteAnsiString(Stream: TStream; const Str: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteAnsiString(Stream: TStream; const Str: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUTF8String_LE(Stream: TStream; const Str: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_WriteUTF8String_BE(Stream: TStream; const Str: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_WriteUTF8String(Stream: TStream; const Str: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUTF8String(Stream: TStream; const Str: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteWideString_LE(Stream: TStream; const Str: WideString; Advance: Boolean = True): TMemSize;

Function Stream_WriteWideString_BE(Stream: TStream; const Str: WideString; Advance: Boolean = True): TMemSize;

Function Stream_WriteWideString(Stream: TStream; const Str: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteWideString(Stream: TStream; const Str: WideString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUnicodeString_LE(Stream: TStream; const Str: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_WriteUnicodeString_BE(Stream: TStream; const Str: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_WriteUnicodeString(Stream: TStream; const Str: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUnicodeString(Stream: TStream; const Str: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteUCS4String_LE(Stream: TStream; const Str: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_WriteUCS4String_BE(Stream: TStream; const Str: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_WriteUCS4String(Stream: TStream; const Str: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUCS4String(Stream: TStream; const Str: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteString_LE(Stream: TStream; const Str: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteString_BE(Stream: TStream; const Str: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteString(Stream: TStream; const Str: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteString(Stream: TStream; const Str: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function Stream_WriteBuffer_LE(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;

Function Stream_WriteBuffer_BE(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;

Function Stream_WriteBuffer(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteBuffer(Stream: TStream; const Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteBytes_LE(Stream: TStream; const Value: array of UInt8; Advance: Boolean = True): TMemSize;

Function Stream_WriteBytes_BE(Stream: TStream; const Value: array of UInt8; Advance: Boolean = True): TMemSize;

Function Stream_WriteBytes(Stream: TStream; const Value: array of UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteBytes(Stream: TStream; const Value: array of UInt8; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_FillBytes_LE(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_FillBytes_BE(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_FillBytes(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_FillBytes(Stream: TStream; Count: TMemSize; Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Stream_WriteVariant_LE(Stream: TStream; const Value: Variant; Advance: Boolean = True): TMemSize;

Function Stream_WriteVariant_BE(Stream: TStream; const Value: Variant; Advance: Boolean = True): TMemSize;

Function Stream_WriteVariant(Stream: TStream; const Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteVariant(Stream: TStream; const Value: Variant; Endian: TEndian = endDefault): TMemSize; overload;


{===============================================================================
--------------------------------------------------------------------------------
                                 Stream reading
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function Stream_ReadBool_LE(Stream: TStream; out Value: ByteBool; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadBool_BE(Stream: TStream; out Value: ByteBool; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadBool(Stream: TStream; out Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadBool(Stream: TStream; out Value: ByteBool; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetBool_LE(Stream: TStream; Advance: Boolean = True): ByteBool;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetBool_BE(Stream: TStream; Advance: Boolean = True): ByteBool;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetBool(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): ByteBool; overload;
Function Stream_GetBool(Stream: TStream; Endian: TEndian = endDefault): ByteBool; overload;

//------------------------------------------------------------------------------

Function Stream_ReadBoolean_LE(Stream: TStream; out Value: Boolean; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadBoolean_BE(Stream: TStream; out Value: Boolean; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadBoolean(Stream: TStream; out Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadBoolean(Stream: TStream; out Value: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetBoolean_LE(Stream: TStream; Advance: Boolean = True): Boolean;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetBoolean_BE(Stream: TStream; Advance: Boolean = True): Boolean;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetBoolean(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetBoolean(Stream: TStream; Endian: TEndian = endDefault): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function Stream_ReadInt8_LE(Stream: TStream; out Value: Int8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadInt8_BE(Stream: TStream; out Value: Int8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadInt8(Stream: TStream; out Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadInt8(Stream: TStream; out Value: Int8; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetInt8_LE(Stream: TStream; Advance: Boolean = True): Int8;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetInt8_BE(Stream: TStream; Advance: Boolean = True): Int8;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetInt8(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int8; overload;
Function Stream_GetInt8(Stream: TStream; Endian: TEndian = endDefault): Int8; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUInt8_LE(Stream: TStream; out Value: UInt8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadUInt8_BE(Stream: TStream; out Value: UInt8; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadUInt8(Stream: TStream; out Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUInt8(Stream: TStream; out Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUInt8_LE(Stream: TStream; Advance: Boolean = True): UInt8;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetUInt8_BE(Stream: TStream; Advance: Boolean = True): UInt8;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetUInt8(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt8; overload;
Function Stream_GetUInt8(Stream: TStream; Endian: TEndian = endDefault): UInt8; overload;

//------------------------------------------------------------------------------

Function Stream_ReadInt16_LE(Stream: TStream; out Value: Int16; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt16_BE(Stream: TStream; out Value: Int16; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt16(Stream: TStream; out Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadInt16(Stream: TStream; out Value: Int16; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetInt16_LE(Stream: TStream; Advance: Boolean = True): Int16;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt16_BE(Stream: TStream; Advance: Boolean = True): Int16;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt16(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int16; overload;
Function Stream_GetInt16(Stream: TStream; Endian: TEndian = endDefault): Int16; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUInt16_LE(Stream: TStream; out Value: UInt16; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt16_BE(Stream: TStream; out Value: UInt16; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt16(Stream: TStream; out Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUInt16(Stream: TStream; out Value: UInt16; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUInt16_LE(Stream: TStream; Advance: Boolean = True): UInt16;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt16_BE(Stream: TStream; Advance: Boolean = True): UInt16;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt16(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt16; overload;
Function Stream_GetUInt16(Stream: TStream; Endian: TEndian = endDefault): UInt16; overload;

//------------------------------------------------------------------------------

Function Stream_ReadInt32_LE(Stream: TStream; out Value: Int32; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt32_BE(Stream: TStream; out Value: Int32; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt32(Stream: TStream; out Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadInt32(Stream: TStream; out Value: Int32; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetInt32_LE(Stream: TStream; Advance: Boolean = True): Int32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt32_BE(Stream: TStream; Advance: Boolean = True): Int32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int32; overload;
Function Stream_GetInt32(Stream: TStream; Endian: TEndian = endDefault): Int32; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUInt32_LE(Stream: TStream; out Value: UInt32; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt32_BE(Stream: TStream; out Value: UInt32; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt32(Stream: TStream; out Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUInt32(Stream: TStream; out Value: UInt32; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUInt32_LE(Stream: TStream; Advance: Boolean = True): UInt32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt32_BE(Stream: TStream; Advance: Boolean = True): UInt32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt32; overload;
Function Stream_GetUInt32(Stream: TStream; Endian: TEndian = endDefault): UInt32; overload;

//------------------------------------------------------------------------------

Function Stream_ReadInt64_LE(Stream: TStream; out Value: Int64; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt64_BE(Stream: TStream; out Value: Int64; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt64(Stream: TStream; out Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadInt64(Stream: TStream; out Value: Int64; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetInt64_LE(Stream: TStream; Advance: Boolean = True): Int64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt64_BE(Stream: TStream; Advance: Boolean = True): Int64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int64; overload;
Function Stream_GetInt64(Stream: TStream; Endian: TEndian = endDefault): Int64; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUInt64_LE(Stream: TStream; out Value: UInt64; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt64_BE(Stream: TStream; out Value: UInt64; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt64(Stream: TStream; out Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUInt64(Stream: TStream; out Value: UInt64; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUInt64_LE(Stream: TStream; Advance: Boolean = True): UInt64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt64_BE(Stream: TStream; Advance: Boolean = True): UInt64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt64; overload;
Function Stream_GetUInt64(Stream: TStream; Endian: TEndian = endDefault): UInt64; overload;

{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Stream_ReadFloat32_LE(Stream: TStream; out Value: Float32; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat32_BE(Stream: TStream; out Value: Float32; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat32(Stream: TStream; out Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadFloat32(Stream: TStream; out Value: Float32; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetFloat32_LE(Stream: TStream; Advance: Boolean = True): Float32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat32_BE(Stream: TStream; Advance: Boolean = True): Float32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float32; overload;
Function Stream_GetFloat32(Stream: TStream; Endian: TEndian = endDefault): Float32; overload;

//------------------------------------------------------------------------------

Function Stream_ReadFloat64_LE(Stream: TStream; out Value: Float64; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat64_BE(Stream: TStream; out Value: Float64; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat64(Stream: TStream; out Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadFloat64(Stream: TStream; out Value: Float64; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetFloat64_LE(Stream: TStream; Advance: Boolean = True): Float64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat64_BE(Stream: TStream; Advance: Boolean = True): Float64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float64; overload;
Function Stream_GetFloat64(Stream: TStream; Endian: TEndian = endDefault): Float64; overload;

//------------------------------------------------------------------------------

Function Stream_ReadFloat80_LE(Stream: TStream; out Value: Float80; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat80_BE(Stream: TStream; out Value: Float80; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat80(Stream: TStream; out Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadFloat80(Stream: TStream; out Value: Float80; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetFloat80_LE(Stream: TStream; Advance: Boolean = True): Float80;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat80_BE(Stream: TStream; Advance: Boolean = True): Float80;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat80(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float80; overload;
Function Stream_GetFloat80(Stream: TStream; Endian: TEndian = endDefault): Float80; overload;

//------------------------------------------------------------------------------

Function Stream_ReadDateTime_LE(Stream: TStream; out Value: TDateTime; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadDateTime_BE(Stream: TStream; out Value: TDateTime; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadDateTime(Stream: TStream; out Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadDateTime(Stream: TStream; out Value: TDateTime; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetDateTime_LE(Stream: TStream; Advance: Boolean = True): TDateTime;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetDateTime_BE(Stream: TStream; Advance: Boolean = True): TDateTime;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetDateTime(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetDateTime(Stream: TStream; Endian: TEndian = endDefault): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function Stream_ReadCurrency_LE(Stream: TStream; out Value: Currency; Advance: Boolean = True): TMemSize;

Function Stream_ReadCurrency_BE(Stream: TStream; out Value: Currency; Advance: Boolean = True): TMemSize;

Function Stream_ReadCurrency(Stream: TStream; out Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadCurrency(Stream: TStream; out Value: Currency; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetCurrency_LE(Stream: TStream; Advance: Boolean = True): Currency;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetCurrency_BE(Stream: TStream; Advance: Boolean = True): Currency;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetCurrency(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Currency; overload;
Function Stream_GetCurrency(Stream: TStream; Endian: TEndian = endDefault): Currency; overload;

{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function Stream_ReadAnsiChar_LE(Stream: TStream; out Value: AnsiChar; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadAnsiChar_BE(Stream: TStream; out Value: AnsiChar; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadAnsiChar(Stream: TStream; out Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadAnsiChar(Stream: TStream; out Value: AnsiChar; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetAnsiChar_LE(Stream: TStream; Advance: Boolean = True): AnsiChar;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetAnsiChar_BE(Stream: TStream; Advance: Boolean = True): AnsiChar;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetAnsiChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): AnsiChar; overload;
Function Stream_GetAnsiChar(Stream: TStream; Endian: TEndian = endDefault): AnsiChar; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUTF8Char_LE(Stream: TStream; out Value: UTF8Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadUTF8Char_BE(Stream: TStream; out Value: UTF8Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadUTF8Char(Stream: TStream; out Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUTF8Char(Stream: TStream; out Value: UTF8Char; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUTF8Char_LE(Stream: TStream; Advance: Boolean = True): UTF8Char;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetUTF8Char_BE(Stream: TStream; Advance: Boolean = True): UTF8Char;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetUTF8Char(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UTF8Char; overload;
Function Stream_GetUTF8Char(Stream: TStream; Endian: TEndian = endDefault): UTF8Char; overload;

//------------------------------------------------------------------------------

Function Stream_ReadWideChar_LE(Stream: TStream; out Value: WideChar; Advance: Boolean = True): TMemSize;

Function Stream_ReadWideChar_BE(Stream: TStream; out Value: WideChar; Advance: Boolean = True): TMemSize;

Function Stream_ReadWideChar(Stream: TStream; out Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadWideChar(Stream: TStream; out Value: WideChar; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetWideChar_LE(Stream: TStream; Advance: Boolean = True): WideChar;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetWideChar_BE(Stream: TStream; Advance: Boolean = True): WideChar;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetWideChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): WideChar; overload;
Function Stream_GetWideChar(Stream: TStream; Endian: TEndian = endDefault): WideChar; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUnicodeChar_LE(Stream: TStream; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;

Function Stream_ReadUnicodeChar_BE(Stream: TStream; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;

Function Stream_ReadUnicodeChar(Stream: TStream; out Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUnicodeChar(Stream: TStream; out Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUnicodeChar_LE(Stream: TStream; Advance: Boolean = True): UnicodeChar;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUnicodeChar_BE(Stream: TStream; Advance: Boolean = True): UnicodeChar;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUnicodeChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UnicodeChar; overload;
Function Stream_GetUnicodeChar(Stream: TStream; Endian: TEndian = endDefault): UnicodeChar; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUCS4Char_LE(Stream: TStream; out Value: UCS4Char; Advance: Boolean = True): TMemSize;

Function Stream_ReadUCS4Char_BE(Stream: TStream; out Value: UCS4Char; Advance: Boolean = True): TMemSize;

Function Stream_ReadUCS4Char(Stream: TStream; out Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUCS4Char(Stream: TStream; out Value: UCS4Char; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUCS4Char_LE(Stream: TStream; Advance: Boolean = True): UCS4Char;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUCS4Char_BE(Stream: TStream; Advance: Boolean = True): UCS4Char;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUCS4Char(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UCS4Char; overload;
Function Stream_GetUCS4Char(Stream: TStream; Endian: TEndian = endDefault): UCS4Char; overload;

//------------------------------------------------------------------------------

Function Stream_ReadChar_LE(Stream: TStream; out Value: Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadChar_BE(Stream: TStream; out Value: Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadChar(Stream: TStream; out Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadChar(Stream: TStream; out Value: Char; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetChar_LE(Stream: TStream; Advance: Boolean = True): Char;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetChar_BE(Stream: TStream; Advance: Boolean = True): Char;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetChar(Stream: TStream; Endian: TEndian = endDefault): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function Stream_ReadShortString_LE(Stream: TStream; out Str: ShortString; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadShortString_BE(Stream: TStream; out Str: ShortString; Advance: Boolean = True): TMemSize;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_ReadShortString(Stream: TStream; out Str: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadShortString(Stream: TStream; out Str: ShortString; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetShortString_LE(Stream: TStream; Advance: Boolean = True): ShortString;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetShortString_BE(Stream: TStream; Advance: Boolean = True): ShortString;{$IFDEF CanInlineFPC} inline;{$ENDIF}

Function Stream_GetShortString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): ShortString; overload;
Function Stream_GetShortString(Stream: TStream; Endian: TEndian = endDefault): ShortString; overload;

//------------------------------------------------------------------------------

Function Stream_ReadAnsiString_LE(Stream: TStream; out Str: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_ReadAnsiString_BE(Stream: TStream; out Str: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_ReadAnsiString(Stream: TStream; out Str: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadAnsiString(Stream: TStream; out Str: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetAnsiString_LE(Stream: TStream; Advance: Boolean = True): AnsiString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetAnsiString_BE(Stream: TStream; Advance: Boolean = True): AnsiString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetAnsiString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): AnsiString; overload;
Function Stream_GetAnsiString(Stream: TStream; Endian: TEndian = endDefault): AnsiString; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUTF8String_LE(Stream: TStream; out Str: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_ReadUTF8String_BE(Stream: TStream; out Str: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_ReadUTF8String(Stream: TStream; out Str: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUTF8String(Stream: TStream; out Str: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUTF8String_LE(Stream: TStream; Advance: Boolean = True): UTF8String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUTF8String_BE(Stream: TStream; Advance: Boolean = True): UTF8String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUTF8String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UTF8String; overload;
Function Stream_GetUTF8String(Stream: TStream; Endian: TEndian = endDefault): UTF8String; overload;

//------------------------------------------------------------------------------

Function Stream_ReadWideString_LE(Stream: TStream; out Str: WideString; Advance: Boolean = True): TMemSize;

Function Stream_ReadWideString_BE(Stream: TStream; out Str: WideString; Advance: Boolean = True): TMemSize;

Function Stream_ReadWideString(Stream: TStream; out Str: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadWideString(Stream: TStream; out Str: WideString; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetWideString_LE(Stream: TStream; Advance: Boolean = True): WideString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetWideString_BE(Stream: TStream; Advance: Boolean = True): WideString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetWideString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): WideString; overload;
Function Stream_GetWideString(Stream: TStream; Endian: TEndian = endDefault): WideString; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUnicodeString_LE(Stream: TStream; out Str: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_ReadUnicodeString_BE(Stream: TStream; out Str: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_ReadUnicodeString(Stream: TStream; out Str: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUnicodeString(Stream: TStream; out Str: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUnicodeString_LE(Stream: TStream; Advance: Boolean = True): UnicodeString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUnicodeString_BE(Stream: TStream; Advance: Boolean = True): UnicodeString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUnicodeString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString; overload;
Function Stream_GetUnicodeString(Stream: TStream; Endian: TEndian = endDefault): UnicodeString; overload;

//------------------------------------------------------------------------------

Function Stream_ReadUCS4String_LE(Stream: TStream; out Str: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_ReadUCS4String_BE(Stream: TStream; out Str: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_ReadUCS4String(Stream: TStream; out Str: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUCS4String(Stream: TStream; out Str: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetUCS4String_LE(Stream: TStream; Advance: Boolean = True): UCS4String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUCS4String_BE(Stream: TStream; Advance: Boolean = True): UCS4String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUCS4String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UCS4String; overload;
Function Stream_GetUCS4String(Stream: TStream; Endian: TEndian = endDefault): UCS4String; overload;

//------------------------------------------------------------------------------

Function Stream_ReadString_LE(Stream: TStream; out Str: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadString_BE(Stream: TStream; out Str: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadString(Stream: TStream; out Str: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadString(Stream: TStream; out Str: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetString_LE(Stream: TStream; Advance: Boolean = True): String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetString_BE(Stream: TStream; Advance: Boolean = True): String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetString(Stream: TStream; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function Stream_ReadBuffer_LE(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;

Function Stream_ReadBuffer_BE(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;

Function Stream_ReadBuffer(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadBuffer(Stream: TStream; out Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Stream_ReadVariant_LE(Stream: TStream; out Value: Variant; Advance: Boolean = True): TMemSize;

Function Stream_ReadVariant_BE(Stream: TStream; out Value: Variant; Advance: Boolean = True): TMemSize;

Function Stream_ReadVariant(Stream: TStream; out Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadVariant(Stream: TStream; out Value: Variant; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetVariant_LE(Stream: TStream; Advance: Boolean = True): Variant;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetVariant_BE(Stream: TStream; Advance: Boolean = True): Variant;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetVariant(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Variant; overload;
Function Stream_GetVariant(Stream: TStream; Endian: TEndian = endDefault): Variant; overload;


(*
{===============================================================================
--------------------------------------------------------------------------------
                                 TCustomStreamer
--------------------------------------------------------------------------------
===============================================================================}
type
  // TValueType is only used internally
  TValueType = (vtShortString,vtAnsiString,vtUTF8String,vtWideString,
                vtUnicodeString,vtUCS4String,vtString,vtFillBytes,vtBytes,
                vtPrimitive1B,vtPrimitive2B,vtPrimitive4B,vtPrimitive8B,
                vtPrimitive10B,vtVariant);

{===============================================================================
    TCustomStreamer - class declaration
===============================================================================}
type
  TCustomStreamer = class(TCustomListObject)
  protected
    fBookmarks:       array of Int64;
    fCount:           Integer;
    fStartPosition:   Int64;
    fEndian:          TEndian;
    Function GetCapacity: Integer; override;
    procedure SetCapacity(Value: Integer); override;
    Function GetCount: Integer; override;
    procedure SetCount(Value: Integer); override;
    Function GetBookmark(Index: Integer): Int64; virtual;
    procedure SetBookmark(Index: Integer; Value: Int64); virtual;
    Function GetCurrentPosition: Int64; virtual; abstract;
    procedure SetCurrentPosition(NewPosition: Int64); virtual; abstract;
    Function GetDistance: Int64; virtual;
    Function WriteValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize; virtual; abstract;
    Function ReadValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize; virtual; abstract;
    procedure Initialize; virtual;
    procedure Finalize; virtual;
  public
    destructor Destroy; override;
    Function LowIndex: Integer; override;
    Function HighIndex: Integer; override;
    procedure MoveToStart; virtual;
    procedure MoveToBookmark(Index: Integer); virtual;
    procedure MoveBy(Offset: Int64); virtual;
    // bookmarks methods
    Function IndexOfBookmark(Position: Int64): Integer; virtual;
    Function AddBookmark: Integer; overload; virtual;
    Function AddBookmark(Position: Int64): Integer; overload; virtual;
    Function RemoveBookmark(Position: Int64; RemoveAll: Boolean = True): Integer; virtual;
    procedure DeleteBookmark(Index: Integer); virtual;
    procedure ClearBookmark; virtual;
    // write methods
    Function WriteBool(Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBoolean(Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt8(Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt8(Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt16(Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt16(Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt32(Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt32(Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt64(Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt64(Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat32(Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat64(Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat80(Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function WriteDateTime(Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function WriteCurrency(Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function WriteAnsiChar(Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUTF8Char(Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteWideChar(Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUnicodeChar(Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUCS4Char(Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteChar(Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteShortString(const Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteAnsiString(const Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUTF8String(const Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteWideString(const Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUnicodeString(const Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUCS4String(const Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteString(const Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBuffer(const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBytes(const Value: array of UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function FillBytes(ByteCount: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteVariant(const Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // read methods
    Function ReadBool(out Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function LoadBool(Advance: Boolean = True): ByteBool; virtual;
    Function ReadBoolean(out Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function LoadBoolean(Advance: Boolean = True): Boolean; virtual;
    Function ReadInt8(out Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function LoadInt8(Advance: Boolean = True): Int8; virtual;
    Function ReadUInt8(out Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUInt8(Advance: Boolean = True): UInt8; virtual;
    Function ReadInt16(out Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function LoadInt16(Advance: Boolean = True): Int16; virtual;
    Function ReadUInt16(out Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUInt16(Advance: Boolean = True): UInt16; virtual;
    Function ReadInt32(out Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function LoadInt32(Advance: Boolean = True): Int32; virtual;
    Function ReadUInt32(out Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUInt32(Advance: Boolean = True): UInt32; virtual;
    Function ReadInt64(out Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function LoadInt64(Advance: Boolean = True): Int64; virtual;
    Function ReadUInt64(out Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUInt64(Advance: Boolean = True): UInt64; virtual;
    Function ReadFloat32(out Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function LoadFloat32(Advance: Boolean = True): Float32; virtual;
    Function ReadFloat64(out Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function LoadFloat64(Advance: Boolean = True): Float64; virtual;
    Function ReadFloat80(out Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function LoadFloat80(Advance: Boolean = True): Float80; virtual;
    Function ReadDateTime(out Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function LoadDateTime(Advance: Boolean = True): TDateTime; virtual;
    Function ReadCurrency(out Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function LoadCurrency(Advance: Boolean = True): Currency; virtual;
    Function ReadAnsiChar(out Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function LoadAnsiChar(Advance: Boolean = True): AnsiChar; virtual;
    Function ReadUTF8Char(out Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUTF8Char(Advance: Boolean = True): UTF8Char; virtual;
    Function ReadWideChar(out Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function LoadWideChar(Advance: Boolean = True): WideChar; virtual;
    Function ReadUnicodeChar(out Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUnicodeChar(Advance: Boolean = True): UnicodeChar; virtual;
    Function ReadUCS4Char(out Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUCS4Char(Advance: Boolean = True): UCS4Char; virtual;
    Function ReadChar(out Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function LoadChar(Advance: Boolean = True): Char; virtual;
    Function ReadShortString(out Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function LoadShortString(Advance: Boolean = True): ShortString; virtual;
    Function ReadAnsiString(out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function LoadAnsiString(Advance: Boolean = True): AnsiString; virtual;
    Function ReadUTF8String(out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUTF8String(Advance: Boolean = True): UTF8String; virtual;
    Function ReadWideString(out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function LoadWideString(Advance: Boolean = True): WideString; virtual;
    Function ReadUnicodeString(out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUnicodeString(Advance: Boolean = True): UnicodeString; virtual;
    Function ReadUCS4String(out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function LoadUCS4String(Advance: Boolean = True): UCS4String; virtual;
    Function ReadString(out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function LoadString(Advance: Boolean = True): String; virtual;
    Function ReadBuffer(var Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function ReadVariant(out Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    Function LoadVariant(Advance: Boolean = True): Variant; virtual;
    // properties
    property Bookmarks[Index: Integer]: Int64 read GetBookmark write SetBookmark;
    property CurrentPosition: Int64 read GetCurrentPosition write SetCurrentPosition;
    property StartPosition: Int64 read fStartPosition;
    property Distance: Int64 read GetDistance;
    property Endian: TEndian read fEndian write fEndian;
  end;

{===============================================================================
--------------------------------------------------------------------------------
                                 TMemoryStreamer
--------------------------------------------------------------------------------
===============================================================================}
{===============================================================================
    TMemoryStreamer - class declaration
===============================================================================}
type
  TMemoryStreamer = class(TCustomStreamer)
  protected
    fCurrentPtr:  Pointer;
    Function GetStartPtr: Pointer; virtual;
    procedure SetBookmark(Index: Integer; Value: Int64); override;
    Function GetCurrentPosition: Int64; override;
    procedure SetCurrentPosition(NewPosition: Int64); override;
    Function WriteValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize; override;
    Function ReadValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize; override;
    procedure Initialize(Memory: Pointer); reintroduce; virtual;
  public
    constructor Create(Memory: Pointer); overload;
    Function IndexOfBookmark(Position: Int64): Integer; override;
    Function AddBookmark(Position: Int64): Integer; override;
    Function RemoveBookmark(Position: Int64; RemoveAll: Boolean = True): Integer; override;
    property CurrentPtr: Pointer read fCurrentPtr write fCurrentPtr;
    property StartPtr: Pointer read GetStartPtr;
  end;

{===============================================================================
--------------------------------------------------------------------------------
                                 TStreamStreamer
--------------------------------------------------------------------------------
===============================================================================}
{===============================================================================
    TStreamStreamer - class declaration
===============================================================================}
type
  TStreamStreamer = class(TCustomStreamer)
  protected
    fTarget:  TStream;
    Function GetCurrentPosition: Int64; override;
    procedure SetCurrentPosition(NewPosition: Int64); override;
    Function WriteValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize; override;
    Function ReadValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize; override;
    procedure Initialize(Target: TStream); reintroduce; virtual;
  public
    constructor Create(Target: TStream);
    property Target: TStream read fTarget;
  end;
*)
implementation

uses
  Variants, Math
  {$IFDEF FPC}, StrRect{$ENDIF};

{$IFDEF FPC_DisableWarns}
  {$DEFINE FPCDWM}
  {$DEFINE W4055:={$WARN 4055 OFF}} // Conversion between ordinals and pointers is not portable
  {$DEFINE W5024:={$WARN 5024 OFF}} // Parameter "$1" not used
  {$DEFINE W5058:={$WARN 5058 OFF}} // Variable "$1" does not seem to be initialized
{$ENDIF}

{$IF SizeOf(UInt32) <> SizeOf(UCS4Char)}  // just to be sure
  {$MESSAGE ERROR 'Type size mismatch (UInt32 - UCS4Char).'}
{$IFEND}

{===============================================================================
    Internal variables
===============================================================================}
var
  ByteOpenArrayIsPacked: Boolean = False;         

{===============================================================================
--------------------------------------------------------------------------------
                               Auxiliary routines
--------------------------------------------------------------------------------
===============================================================================}

procedure AdvancePointer(Advance: Boolean; var Ptr: Pointer; Offset: TMemSize);{$IFDEF CanInline} inline;{$ENDIF}
begin
{$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
If Advance then
  Ptr := Pointer(TMemSize(PtrUInt(Ptr)) + Offset);
{$IFDEF FPCDWM}{$POP}{$ENDIF}
end;

//------------------------------------------------------------------------------

procedure AdvanceStream(Advance: Boolean; Stream: TStream; Offset: TMemSize);{$IFDEF CanInline} inline;{$ENDIF}
begin
If not Advance then
  Stream.Seek(-Int64(Offset),soCurrent);
end;

//==============================================================================

Function BoolToNum(Val: Boolean): UInt8;{$IFDEF CanInline} inline;{$ENDIF}
begin
If Val then Result := $FF
  else Result := 0;
end;

//------------------------------------------------------------------------------

Function NumToBool(Val: UInt8): Boolean;{$IFDEF CanInline} inline;{$ENDIF}
begin
Result := Val <> 0;
end;

//------------------------------------------------------------------------------

procedure ClampStringLength(var StrLength: Int32);{$IFDEF CanInline} inline;{$ENDIF}
begin
If StrLength < 0 then
  StrLength := 0;
end;

//==============================================================================
type
  Int32Rec = packed record
    LoWord: UInt16;
    HiWord: UInt16;
  end;

//------------------------------------------------------------------------------

Function SwapEndian(Value: UInt16): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}
begin
Result := UInt16(Value shl 8) or UInt16(Value shr 8);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function SwapEndian(Value: UInt32): UInt32; overload;
begin
Int32Rec(Result).HiWord := SwapEndian(Int32Rec(Value).LoWord);
Int32Rec(Result).LoWord := SwapEndian(Int32Rec(Value).HiWord);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function SwapEndian(Value: UInt64): UInt64; overload;
begin
Int64Rec(Result).Hi := SwapEndian(Int64Rec(Value).Lo);
Int64Rec(Result).Lo := SwapEndian(Int64Rec(Value).Hi);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

type
  TFloat80Overlay = packed array[0..9] of Byte;

Function SwapEndian(Value: TFloat80Overlay): TFloat80Overlay; overload;
begin
Result[0] := Value[9];
Result[1] := Value[8];
Result[2] := Value[7];
Result[3] := Value[6];
Result[4] := Value[5];
Result[5] := Value[4];
Result[6] := Value[3];
Result[7] := Value[2];
Result[8] := Value[1];
Result[9] := Value[0];
end;

//==============================================================================

Function Ptr_WriteUInt16Arr_SwapEndian(var Dest: PUInt16; Data: PUInt16; Length: TStrSize): TMemSize;
var
  i:  TStrSize;
begin
Result := 0;
For i := 1 to Length do
  begin
    Dest^ := SwapEndian(Data^);
    Inc(Dest);
    Inc(Data);
    Inc(Result,SizeOf(UInt16));
  end;
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt32Arr_SwapEndian(var Dest: PUInt32; Data: PUInt32; Length: TStrSize): TMemSize;
var
  i:  TStrSize;
begin
Result := 0;
For i := 1 to Length do
  begin
    Dest^ := SwapEndian(Data^);
    Inc(Dest);
    Inc(Data);
    Inc(Result,SizeOf(UInt32));
  end;
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt16Arr_SwapEndian(var Src: PUInt16; Data: PUInt16; Length: TStrSize): TMemSize;
var
  i:  TStrSize;
begin
Result := 0;
For i := 1 to Length do
  begin
    Data^ := SwapEndian(Src^);
    Inc(Src);
    Inc(Data);
    Inc(Result,SizeOf(UInt16));
  end;
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt32Arr_SwapEndian(var Src: PUInt32; Data: PUInt32; Length: TStrSize): TMemSize;
var
  i:  TStrSize;
begin
Result := 0;
For i := 1 to Length do
  begin
    Data^ := SwapEndian(Src^);
    Inc(Src);
    Inc(Data);
    Inc(Result,SizeOf(UInt32));
  end;
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt16Arr_SwapEndian(Stream: TStream; Data: PUInt16; Length: TStrSize): TMemSize;
var
  Buffer:   array[0..511] of UInt16;
  CopyCnt:  TStrSize;
  i:        TStrSize;
begin
{
  Buffering might not be needed when operating in TCustomMemoryStream and its
  descendants, but it should not hurt.
}
Result := 0;
while Length > 0 do
  begin
    CopyCnt := Min(Length,System.Length(Buffer));
    For i := 0 to Pred(CopyCnt) do
      begin
        Buffer[i] := SwapEndian(Data^);
        Inc(Data);
        Inc(Result,SizeOf(UInt16));
      end;
    Stream.WriteBuffer(Addr(Buffer)^,CopyCnt * SizeOf(UInt16));
    Dec(Length,CopyCnt);
  end;
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt32Arr_SwapEndian(Stream: TStream; Data: PUInt32; Length: TStrSize): TMemSize;
var
  Buffer:   array[0..255] of UInt32;
  CopyCnt:  TStrSize;
  i:        TStrSize;
begin
Result := 0;
while Length > 0 do
  begin
    CopyCnt := Min(Length,System.Length(Buffer));
    For i := 0 to Pred(CopyCnt) do
      begin
        Buffer[i] := SwapEndian(Data^);
        Inc(Data);
        Inc(Result,SizeOf(UInt32));
      end;
    Stream.WriteBuffer(Addr(Buffer)^,CopyCnt * SizeOf(UInt32));
    Dec(Length,CopyCnt);
  end;
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt16Arr_SwapEndian(Stream: TStream; Data: PUInt16; Length: TStrSize): TMemSize;
var
  Buffer:   array[0..511] of UInt16;
  CopyCnt:  TStrSize;
  i:        TStrSize;
begin
Result := 0;
while Length > 0 do
  begin
    CopyCnt := Min(Length,System.Length(Buffer));
    Stream.ReadBuffer(Addr(Buffer)^,CopyCnt * SizeOf(UInt16));
    For i := 0 to Pred(CopyCnt) do
      begin
        Data^ := SwapEndian(Buffer[i]);
        Inc(Data);
        Inc(Result,SizeOf(UInt16));
      end;
    Dec(Length,CopyCnt);
  end;
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt32Arr_SwapEndian(Stream: TStream; Data: PUInt32; Length: TStrSize): TMemSize;
var
  Buffer:   array[0..255] of UInt32;
  CopyCnt:  TStrSize;
  i:        TStrSize;
begin
Result := 0;
while Length > 0 do
  begin
    CopyCnt := Min(Length,System.Length(Buffer));
    Stream.ReadBuffer(Addr(Buffer)^,CopyCnt * SizeOf(UInt32));
    For i := 0 to Pred(CopyCnt) do
      begin
        Data^ := SwapEndian(Buffer[i]);
        Inc(Data);
        Inc(Result,SizeOf(UInt32));
      end;
    Dec(Length,CopyCnt);
  end;
end;


{===============================================================================
--------------------------------------------------------------------------------
                                Variant utilities
--------------------------------------------------------------------------------
===============================================================================}
{
  Following constants and functions are here to strictly separate implementation
  details from the actual streaming.
}
const
  BS_VARTYPENUM_BOOLEN   = 0;
  BS_VARTYPENUM_SHORTINT = 1;
  BS_VARTYPENUM_SMALLINT = 2;
  BS_VARTYPENUM_INTEGER  = 3;
  BS_VARTYPENUM_INT64    = 4;
  BS_VARTYPENUM_BYTE     = 5;
  BS_VARTYPENUM_WORD     = 6;
  BS_VARTYPENUM_LONGWORD = 7;
  BS_VARTYPENUM_UINT64   = 8;
  BS_VARTYPENUM_SINGLE   = 9;
  BS_VARTYPENUM_DOUBLE   = 10;
  BS_VARTYPENUM_CURRENCY = 11;
  BS_VARTYPENUM_DATE     = 12;
  BS_VARTYPENUM_OLESTR   = 13;
  BS_VARTYPENUM_STRING   = 14;
  BS_VARTYPENUM_USTRING  = 15;
  BS_VARTYPENUM_VARIANT  = 16;

  BS_VARTYPENUM_TYPEMASK = $3F;
  BS_VARTYPENUM_ARRAY    = $80;

//------------------------------------------------------------------------------

Function VarTypeToInt(VarType: TVarType): UInt8;
begin
// varByRef is ignored
case VarType and varTypeMask of
  varBoolean:   Result := BS_VARTYPENUM_BOOLEN;
  varShortInt:  Result := BS_VARTYPENUM_SHORTINT;
  varSmallint:  Result := BS_VARTYPENUM_SMALLINT;
  varInteger:   Result := BS_VARTYPENUM_INTEGER;
  varInt64:     Result := BS_VARTYPENUM_INT64;
  varByte:      Result := BS_VARTYPENUM_BYTE;
  varWord:      Result := BS_VARTYPENUM_WORD;
  varLongWord:  Result := BS_VARTYPENUM_LONGWORD;
{$IF Declared(varUInt64)}
  varUInt64:    Result := BS_VARTYPENUM_UINT64;
{$IFEND}
  varSingle:    Result := BS_VARTYPENUM_SINGLE;
  varDouble:    Result := BS_VARTYPENUM_DOUBLE;
  varCurrency:  Result := BS_VARTYPENUM_CURRENCY;
  varDate:      Result := BS_VARTYPENUM_DATE;
  varOleStr:    REsult := BS_VARTYPENUM_OLESTR;
  varString:    Result := BS_VARTYPENUM_STRING;
{$IF Declared(varUString)}
  varUString:   Result := BS_VARTYPENUM_USTRING;
{$IFEND}
  varVariant:   Result := BS_VARTYPENUM_VARIANT;
else
  raise EBSUnsupportedVarType.CreateFmt('VarTypeToInt: Unsupported variant type (%d).',[VarType{do not mask the type}]);
end;
If VarType and varArray <> 0 then
  Result := Result or BS_VARTYPENUM_ARRAY;
end;

//------------------------------------------------------------------------------

Function IntToVarType(Value: UInt8): TVarType;
begin
case Value and BS_VARTYPENUM_TYPEMASK of
  BS_VARTYPENUM_BOOLEN:   Result := varBoolean;
  BS_VARTYPENUM_SHORTINT: Result := varShortInt;
  BS_VARTYPENUM_SMALLINT: Result := varSmallint;
  BS_VARTYPENUM_INTEGER:  Result := varInteger;
  BS_VARTYPENUM_INT64:    Result := varInt64;
  BS_VARTYPENUM_BYTE:     Result := varByte;
  BS_VARTYPENUM_WORD:     Result := varWord;
  BS_VARTYPENUM_LONGWORD: Result := varLongWord;
  BS_VARTYPENUM_UINT64:   Result := {$IF Declared(varUInt64)}varUInt64{$ELSE}varInt64{$IFEND};
  BS_VARTYPENUM_SINGLE:   Result := varSingle;
  BS_VARTYPENUM_DOUBLE:   Result := varDouble;
  BS_VARTYPENUM_CURRENCY: Result := varCurrency;
  BS_VARTYPENUM_DATE:     Result := varDate;
  BS_VARTYPENUM_OLESTR:   Result := varOleStr;
  BS_VARTYPENUM_STRING:   Result := varString;
  BS_VARTYPENUM_USTRING:  Result := {$IF Declared(varUString)}varUString{$ELSE}varOleStr{$IFEND};
  BS_VARTYPENUM_VARIANT:  Result := varVariant;
else
  raise EBSUnsupportedVarType.CreateFmt('IntToVarType: Unsupported variant type number (%d).',[Value]);
end;
If Value and BS_VARTYPENUM_ARRAY <> 0 then
  Result := Result or varArray;
end;

//------------------------------------------------------------------------------

Function VarToBool(const Value: Variant): Boolean;
begin
Result := Value;
end;

//------------------------------------------------------------------------------

Function VarToInt64(const Value: Variant): Int64;
begin
Result := Value;
end;

//------------------------------------------------------------------------------

Function VarToUInt64(const Value: Variant): UInt64;
begin
Result := Value;
end;

{===============================================================================
    Endianess - implementation
===============================================================================}

Function ResolveEndian(Endian: TEndian): TEndian;
begin
case Endian of
  endSystem:  Result := SysEndian;
  endDefault: Result := endLittle;  
else
  Result := Endian;
end;
end;


{===============================================================================
--------------------------------------------------------------------------------
                               Allocation helpers
--------------------------------------------------------------------------------
===============================================================================}

Function StreamedSize_Bool: TMemSize;
begin
Result := SizeOf(UInt8);
end;

//------------------------------------------------------------------------------

Function StreamedSize_Boolean: TMemSize;
begin
Result := StreamedSize_Bool;
end;

//==============================================================================

Function StreamedSize_Int8: TMemSize;
begin
Result := SizeOf(Int8);
end;

//------------------------------------------------------------------------------

Function StreamedSize_UInt8: TMemSize;
begin
Result := SizeOf(UInt8);
end;

//------------------------------------------------------------------------------

Function StreamedSize_Int16: TMemSize;
begin
Result := SizeOf(Int16);
end;

//------------------------------------------------------------------------------

Function StreamedSize_UInt16: TMemSize;
begin
Result := SizeOf(UInt16);
end;

//------------------------------------------------------------------------------

Function StreamedSize_Int32: TMemSize;
begin
Result := SizeOf(Int32);
end;

//------------------------------------------------------------------------------

Function StreamedSize_UInt32: TMemSize;
begin
Result := SizeOf(UInt32);
end;

//------------------------------------------------------------------------------

Function StreamedSize_Int64: TMemSize;
begin
Result := SizeOf(Int64);
end;

//------------------------------------------------------------------------------

Function StreamedSize_UInt64: TMemSize;
begin
Result := SizeOf(UInt64);
end;

//==============================================================================

Function StreamedSize_Float32: TMemSize;
begin
Result := SizeOf(Float32);
end;

//------------------------------------------------------------------------------

Function StreamedSize_Float64: TMemSize;
begin
Result := SizeOf(Float64);
end;

//------------------------------------------------------------------------------

Function StreamedSize_Float80: TMemSize;
begin
Result := SizeOf(Float80);
end;

//------------------------------------------------------------------------------

Function StreamedSize_DateTime: TMemSize;
begin
Result := StreamedSize_Float64;
end;

//------------------------------------------------------------------------------

Function StreamedSize_Currency: TMemSize;
begin
Result := SizeOf(Currency);
end;

//==============================================================================

Function StreamedSize_AnsiChar: TMemSize;
begin
Result := SizeOf(AnsiChar);
end;

//------------------------------------------------------------------------------

Function StreamedSize_UTF8Char: TMemSize;
begin
Result := SizeOf(UTF8Char);
end;

//------------------------------------------------------------------------------

Function StreamedSize_WideChar: TMemSize;
begin
Result := SizeOf(WideChar);
end;

//------------------------------------------------------------------------------

Function StreamedSize_UnicodeChar: TMemSize;
begin
Result := SizeOf(UnicodeChar);
end;

//------------------------------------------------------------------------------

Function StreamedSize_UCS4Char: TMemSize;
begin
Result := SizeOf(UCS4Char);
end;

//------------------------------------------------------------------------------

Function StreamedSize_Char: TMemSize;
begin
Result := StreamedSize_UInt16;
end;

//==============================================================================

Function StreamedSize_ShortString(const Str: ShortString): TMemSize;
begin
Result := StreamedSize_UInt8 + TMemSize(Length(Str));
end;     

//------------------------------------------------------------------------------

Function StreamedSize_AnsiString(const Str: AnsiString): TMemSize;
begin
Result := StreamedSize_Int32 + TMemSize(Length(Str) * SizeOf(AnsiChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_UTF8String(const Str: UTF8String): TMemSize;
begin
Result := StreamedSize_Int32 + TMemSize(Length(Str) * SizeOf(UTF8Char));
end;

//------------------------------------------------------------------------------

Function StreamedSize_WideString(const Str: WideString): TMemSize;
begin
Result := StreamedSize_Int32 + TMemSize(Length(Str) * SizeOf(WideChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_UnicodeString(const Str: UnicodeString): TMemSize;
begin
Result := StreamedSize_Int32 + TMemSize(Length(Str) * SizeOf(UnicodeChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_UCS4String(const Str: UCS4String): TMemSize;
begin
{
  Note that UCS4 string CAN (we must cope with situation it does not) contain
  an explicit terminating zero, which is not being streamed.
}
If Length(Str) > 0 then
  Result := StreamedSize_Int32 + TMemSize(IfThen(Str[High(Str)] = 0,Pred(Length(Str)),Length(Str)) * SizeOf(UCS4Char))
else
  Result := StreamedSize_Int32;
end;

//------------------------------------------------------------------------------

Function StreamedSize_String(const Str: String): TMemSize;
begin
Result := StreamedSize_UTF8String(StrToUTF8(Str));
end;

//==============================================================================

Function StreamedSize_Buffer(Size: TMemSize): TMemSize;
begin
Result := Size;
end;

//------------------------------------------------------------------------------

Function StreamedSize_Bytes(Count: TMemSize): TMemSize;
begin
Result := Count;
end;

//==============================================================================

Function StreamedSize_Variant(const Value: Variant): TMemSize;
var
  Dimensions: Integer;
  Indices:    array of Integer;

  Function StreamedSize_VarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:  Integer;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          Inc(Result,StreamedSize_Variant(VarArrayGet(Value,Indices)))
        else
          Inc(Result,StreamedSize_VarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
If VarIsArray(Value) then
  begin
    // array, need to do complete scan to get full size
    Dimensions := VarArrayDimCount(Value);
    Result := 5{type (byte) + number of dimensions (int32)} +
              (Dimensions * 8{2 * int32 (indices bounds)});
    If Dimensions > 0 then
      begin
        SetLength(Indices,Dimensions);
        Inc(Result,StreamedSize_VarArrayDimension(1));
      end;
  end
else
  begin
    // simple type
    If VarType(Value) <> (varVariant or varByRef) then
      begin
        case VarType(Value) and varTypeMask of  
          varBoolean:   Result := StreamedSize_Boolean;
          varShortInt:  Result := StreamedSize_Int8;
          varSmallint:  Result := StreamedSize_Int16;
          varInteger:   Result := StreamedSize_Int32;
          varInt64:     Result := StreamedSize_Int64;
          varByte:      Result := StreamedSize_UInt8;
          varWord:      Result := StreamedSize_UInt16;
          varLongWord:  Result := StreamedSize_UInt32;
        {$IF Declared(varUInt64)}
          varUInt64:    Result := StreamedSize_UInt64;
        {$IFEND}
          varSingle:    Result := StreamedSize_Float32;
          varDouble:    Result := StreamedSize_Float64;
          varCurrency:  Result := StreamedSize_Currency;
          varDate:      Result := StreamedSize_DateTime;
          varOleStr:    Result := StreamedSize_WideString(Value);
          varString:    Result := StreamedSize_AnsiString(AnsiString(Value));
        {$IF Declared(varUInt64)}
          varUString:   Result := StreamedSize_UnicodeString(Value);
        {$IFEND}
        else
          raise EBSUnsupportedVarType.CreateFmt('StreamedSize_Variant: Cannot scan variant of this type (%d).',[VarType(Value) and varTypeMask]);
        end;
        Inc(Result);  // for the variant type byte
      end
    else Result := StreamedSize_Variant(Variant(FindVarData(Value)^));
  end;
end;


{===============================================================================
--------------------------------------------------------------------------------
                                 Memory writing
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function _Ptr_WriteBool(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize;
begin
UInt8(Dest^) := BoolToNum(Value);
Result := SizeOf(UInt8);
AdvancePointer(Advance,Dest,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBool_LE(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBool(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteBool_LE(Dest: Pointer; Value: ByteBool): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteBool(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBool_BE(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBool(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteBool_BE(Dest: Pointer; Value: ByteBool): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteBool(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBool(var Dest: Pointer; Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteBool_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteBool_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteBool(Dest: Pointer; Value: ByteBool; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteBool_BE(Ptr,Value,False)
else
  Result := Ptr_WriteBool_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteBoolean_LE(var Dest: Pointer; Value: Boolean; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteBool_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteBoolean_LE(Dest: Pointer; Value: Boolean): TMemSize;
begin
Result := Ptr_WriteBool_LE(Dest,Value);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBoolean_BE(var Dest: Pointer; Value: Boolean; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteBool_BE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteBoolean_BE(Dest: Pointer; Value: Boolean): TMemSize;
begin
Result := Ptr_WriteBool_BE(Dest,Value);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBoolean(var Dest: Pointer; Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteBool(Dest,Value,Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteBoolean(Dest: Pointer; Value: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteBool(Dest,Value,Endian);
end;

{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function _Ptr_WriteInt8(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize;
begin
Int8(Dest^) := Value;
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteInt8_LE(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteInt8(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt8_LE(Dest: Pointer; Value: Int8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteInt8(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteInt8_BE(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteInt8(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt8_BE(Dest: Pointer; Value: Int8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteInt8(Ptr,Value,False);
end;     

//------------------------------------------------------------------------------

Function Ptr_WriteInt8(var Dest: Pointer; Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteInt8_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteInt8_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt8(Dest: Pointer; Value: Int8; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteInt8_BE(Ptr,Value,False)
else
  Result := Ptr_WriteInt8_LE(Ptr,Value,False);
end;

//==============================================================================

Function _Ptr_WriteUInt8(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize;
begin
UInt8(Dest^) := Value;
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt8_LE(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteUInt8(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt8_LE(Dest: Pointer; Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteUInt8(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt8_BE(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteUInt8(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt8_BE(Dest: Pointer; Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteUInt8(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt8(var Dest: Pointer; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUInt8_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUInt8_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt8(Dest: Pointer; Value: UInt8; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUInt8_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUInt8_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteInt16_LE(var Dest: Pointer; Value: Int16; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Int16(Dest^) := Int16(SwapEndian(UInt16(Value)));
{$ELSE}
Int16(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt16_LE(Dest: Pointer; Value: Int16): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteInt16_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteInt16_BE(var Dest: Pointer; Value: Int16; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Int16(Dest^) := Value;
{$ELSE}
Int16(Dest^) := Int16(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt16_BE(Dest: Pointer; Value: Int16): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteInt16_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteInt16(var Dest: Pointer; Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteInt16_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteInt16_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt16(Dest: Pointer; Value: Int16; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteInt16_BE(Ptr,Value,False)
else
  Result := Ptr_WriteInt16_LE(Ptr,Value,False);
end;
 
//==============================================================================

Function Ptr_WriteUInt16_LE(var Dest: Pointer; Value: UInt16; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt16(Dest^) := SwapEndian(Value);
{$ELSE}
UInt16(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt16_LE(Dest: Pointer; Value: UInt16): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUInt16_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt16_BE(var Dest: Pointer; Value: UInt16; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt16(Dest^) := Value;
{$ELSE}
UInt16(Dest^) := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt16_BE(Dest: Pointer; Value: UInt16): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUInt16_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt16(var Dest: Pointer; Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUInt16_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUInt16_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt16(Dest: Pointer; Value: UInt16; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUInt16_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUInt16_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteInt32_LE(var Dest: Pointer; Value: Int32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Int32(Dest^) := Int32(SwapEndian(UInt32(Value)));
{$ELSE}
Int32(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt32_LE(Dest: Pointer; Value: Int32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteInt32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteInt32_BE(var Dest: Pointer; Value: Int32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Int32(Dest^) := Value;
{$ELSE}
Int32(Dest^) := Int32(SwapEndian(UInt32(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt32_BE(Dest: Pointer; Value: Int32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteInt32_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteInt32(var Dest: Pointer; Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteInt32_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteInt32_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt32(Dest: Pointer; Value: Int32; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteInt32_BE(Ptr,Value,False)
else
  Result := Ptr_WriteInt32_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteUInt32_LE(var Dest: Pointer; Value: UInt32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt32(Dest^) := SwapEndian(Value);
{$ELSE}
UInt32(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt32_LE(Dest: Pointer; Value: UInt32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUInt32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt32_BE(var Dest: Pointer; Value: UInt32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt32(Dest^) := Value;
{$ELSE}
UInt32(Dest^) := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt32_BE(Dest: Pointer; Value: UInt32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUInt32_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt32(var Dest: Pointer; Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUInt32_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUInt32_LE(Dest,Value,Advance);
end;
 
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt32(Dest: Pointer; Value: UInt32; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUInt32_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUInt32_LE(Ptr,Value,False);
end;
  
//==============================================================================

Function Ptr_WriteInt64_LE(var Dest: Pointer; Value: Int64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Int64(Dest^) := Int64(SwapEndian(UInt64(Value)));
{$ELSE}
Int64(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt64_LE(Dest: Pointer; Value: Int64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteInt64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteInt64_BE(var Dest: Pointer; Value: Int64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Int64(Dest^) := Value;
{$ELSE}
Int64(Dest^) := Int64(SwapEndian(UInt64(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt64_BE(Dest: Pointer; Value: Int64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteInt64_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteInt64(var Dest: Pointer; Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteInt64_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteInt64_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteInt64(Dest: Pointer; Value: Int64; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteInt64_BE(Ptr,Value,False)
else
  Result := Ptr_WriteInt64_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteUInt64_LE(var Dest: Pointer; Value: UInt64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt64(Dest^) := SwapEndian(Value);
{$ELSE}
UInt64(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt64_LE(Dest: Pointer; Value: UInt64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUInt64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt64_BE(var Dest: Pointer; Value: UInt64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt64(Dest^) := Value;
{$ELSE}
UInt64(Dest^) := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt64_BE(Dest: Pointer; Value: UInt64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUInt64_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUInt64(var Dest: Pointer; Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUInt64_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUInt64_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUInt64(Dest: Pointer; Value: UInt64; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUInt64_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUInt64_LE(Ptr,Value,False);
end;

{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Ptr_WriteFloat32_LE(var Dest: Pointer; Value: Float32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
{
  Casting to UInt32 is done to prevent invalid floating point operation when
  assigning byte-swapped value, which can be SNaN or denormal.
}
UInt32(Dest^) := SwapEndian(UInt32(Addr(Value)^));
{$ELSE}
Float32(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat32_LE(Dest: Pointer; Value: Float32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteFloat32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteFloat32_BE(var Dest: Pointer; Value: Float32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Float32(Dest^) := Value;
{$ELSE}
UInt32(Dest^) := SwapEndian(UInt32(Addr(Value)^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat32_BE(Dest: Pointer; Value: Float32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteFloat32_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteFloat32(var Dest: Pointer; Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteFloat32_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteFloat32_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat32(Dest: Pointer; Value: Float32; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteFloat32_BE(Ptr,Value,False)
else
  Result := Ptr_WriteFloat32_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteFloat64_LE(var Dest: Pointer; Value: Float64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt64(Dest^) := SwapEndian(UInt64(Addr(Value)^));
{$ELSE}
Float64(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat64_LE(Dest: Pointer; Value: Float64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteFloat64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteFloat64_BE(var Dest: Pointer; Value: Float64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Float64(Dest^) := Value;
{$ELSE}
UInt64(Dest^) := SwapEndian(UInt64(Addr(Value)^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat64_BE(Dest: Pointer; Value: Float64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteFloat64_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteFloat64(var Dest: Pointer; Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteFloat64_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteFloat64_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat64(Dest: Pointer; Value: Float64; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteFloat64_BE(Ptr,Value,False)
else
  Result := Ptr_WriteFloat64_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteFloat80_LE(var Dest: Pointer; Value: Float80; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
TFloat80Overlay(Dest^) := SwapEndian(TFloat80Overlay(Value));
{$ELSE}
Float80(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat80_LE(Dest: Pointer; Value: Float80): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteFloat80_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteFloat80_BE(var Dest: Pointer; Value: Float80; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Float80(Dest^) := Value;
{$ELSE}
TFloat80Overlay(Dest^) := SwapEndian(TFloat80Overlay(Value));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat80_BE(Dest: Pointer; Value: Float80): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteFloat80_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteFloat80(var Dest: Pointer; Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteFloat80_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteFloat80_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteFloat80(Dest: Pointer; Value: Float80; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteFloat80_BE(Ptr,Value,False)
else
  Result := Ptr_WriteFloat80_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteDateTime_LE(var Dest: Pointer; Value: TDateTime; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteFloat64_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteDateTime_LE(Dest: Pointer; Value: TDateTime): TMemSize;
begin
Result := Ptr_WriteFloat64_LE(Dest,Value);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteDateTime_BE(var Dest: Pointer; Value: TDateTime; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteFloat64_BE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteDateTime_BE(Dest: Pointer; Value: TDateTime): TMemSize;
begin
Result := Ptr_WriteFloat64_BE(Dest,Value);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteDateTime(var Dest: Pointer; Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteFloat64(Dest,Value,Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteDateTime(Dest: Pointer; Value: TDateTime; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteFloat64(Dest,Value,Endian);
end;

//==============================================================================

Function Ptr_WriteCurrency_LE(var Dest: Pointer; Value: Currency; Advance: Boolean): TMemSize;
begin
// prevent implicit conversion of currency to other types
{$IFDEF ENDIAN_BIG}
UInt64(Dest^) := SwapEndian(UInt64(Addr(Value)^));
{$ELSE}
UInt64(Dest^) := UInt64(Addr(Value)^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteCurrency_LE(Dest: Pointer; Value: Currency): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteCurrency_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteCurrency_BE(var Dest: Pointer; Value: Currency; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt64(Dest^) := UInt64(Addr(Value)^);
{$ELSE}
UInt64(Dest^) := SwapEndian(UInt64(Addr(Value)^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteCurrency_BE(Dest: Pointer; Value: Currency): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteCurrency_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteCurrency(var Dest: Pointer; Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteCurrency_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteCurrency_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteCurrency(Dest: Pointer; Value: Currency; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteCurrency_BE(Ptr,Value,False)
else
  Result := Ptr_WriteCurrency_LE(Ptr,Value,False);
end;

{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function _Ptr_WriteAnsiChar(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize;
begin
AnsiChar(Dest^) := Value;
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiChar_LE(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteAnsiChar(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiChar_LE(Dest: Pointer; Value: AnsiChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteAnsiChar(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiChar_BE(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteAnsiChar(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiChar_BE(Dest: Pointer; Value: AnsiChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteAnsiChar(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiChar(var Dest: Pointer; Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteAnsiChar_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteAnsiChar_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiChar(Dest: Pointer; Value: AnsiChar; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteAnsiChar_BE(Ptr,Value,False)
else
  Result := Ptr_WriteAnsiChar_LE(Ptr,Value,False);
end;

//==============================================================================

Function _Ptr_WriteUTF8Char(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize;
begin
UTF8Char(Dest^) := Value;
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8Char_LE(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteUTF8Char(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUTF8Char_LE(Dest: Pointer; Value: UTF8Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteUTF8Char(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8Char_BE(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteUTF8Char(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUTF8Char_BE(Dest: Pointer; Value: UTF8Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteUTF8Char(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8Char(var Dest: Pointer; Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUTF8Char_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUTF8Char_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUTF8Char(Dest: Pointer; Value: UTF8Char; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUTF8Char_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUTF8Char_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteWideChar_LE(var Dest: Pointer; Value: WideChar; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
WideChar(Dest^) := WideChar(SwapEndian(UInt16(Value)));
{$ELSE}
WideChar(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideChar_LE(Dest: Pointer; Value: WideChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteWideChar_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteWideChar_BE(var Dest: Pointer; Value: WideChar; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
WideChar(Dest^) := Value;
{$ELSE}
WideChar(Dest^) := WideChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideChar_BE(Dest: Pointer; Value: WideChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteWideChar_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteWideChar(var Dest: Pointer; Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteWideChar_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteWideChar_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideChar(Dest: Pointer; Value: WideChar; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteWideChar_BE(Ptr,Value,False)
else
  Result := Ptr_WriteWideChar_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteUnicodeChar_LE(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UnicodeChar(Dest^) := UnicodeChar(SwapEndian(UInt16(Value)));
{$ELSE}
UnicodeChar(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeChar_LE(Dest: Pointer; Value: UnicodeChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUnicodeChar_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeChar_BE(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UnicodeChar(Dest^) := Value;
{$ELSE}
UnicodeChar(Dest^) := UnicodeChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeChar_BE(Dest: Pointer; Value: UnicodeChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUnicodeChar_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeChar(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUnicodeChar_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUnicodeChar_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeChar(Dest: Pointer; Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUnicodeChar_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUnicodeChar_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteUCS4Char_LE(var Dest: Pointer; Value: UCS4Char; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt32(Dest^) := SwapEndian(UInt32(Value));
{$ELSE}
UCS4Char(Dest^) := Value;
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4Char_LE(Dest: Pointer; Value: UCS4Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUCS4Char_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUCS4Char_BE(var Dest: Pointer; Value: UCS4Char; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UCS4Char(Dest^) := Value;
{$ELSE}
UInt32(Dest^) := SwapEndian(UInt32(Value));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4Char_BE(Dest: Pointer; Value: UCS4Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUCS4Char_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUCS4Char(var Dest: Pointer; Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUCS4Char_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUCS4Char_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4Char(Dest: Pointer; Value: UCS4Char; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUCS4Char_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUCS4Char_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteChar_LE(var Dest: Pointer; Value: Char; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteUInt16_LE(Dest,UInt16(Ord(Value)),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteChar_LE(Dest: Pointer; Value: Char): TMemSize;
begin
Result := Ptr_WriteUInt16_LE(Dest,UInt16(Ord(Value)));
end;

//------------------------------------------------------------------------------

Function Ptr_WriteChar_BE(var Dest: Pointer; Value: Char; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteUInt16_BE(Dest,UInt16(Ord(Value)),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteChar_BE(Dest: Pointer; Value: Char): TMemSize;
begin
Result := Ptr_WriteUInt16_BE(Dest,UInt16(Ord(Value)));
end;

//------------------------------------------------------------------------------

Function Ptr_WriteChar(var Dest: Pointer; Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteUInt16(Dest,UInt16(Ord(Value)),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteChar(Dest: Pointer; Value: Char; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteUInt16(Dest,UInt16(Ord(Value)),Endian);
end;

{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function _Ptr_WriteShortString(var Dest: Pointer; const Str: ShortString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteUInt8_LE(WorkPtr,UInt8(Length(Str)),True);
If Length(Str) > 0 then
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,(Addr(Str[1]))^,Length(Str),True));
If Advance then
  Dest := WorkPtr;
end;

//------------------------------------------------------------------------------

Function Ptr_WriteShortString_LE(var Dest: Pointer; const Str: ShortString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteShortString(Dest,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteShortString_LE(Dest: Pointer; const Str: ShortString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteShortString(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteShortString_BE(var Dest: Pointer; const Str: ShortString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteShortString(Dest,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteShortString_BE(Dest: Pointer; const Str: ShortString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteShortString(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteShortString(var Dest: Pointer; const Str: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteShortString_BE(Dest,Str,Advance)
else
  Result := Ptr_WriteShortString_LE(Dest,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteShortString(Dest: Pointer; const Str: ShortString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteShortString_BE(Ptr,Str,False)
else
  Result := Ptr_WriteShortString_LE(Ptr,Str,False);
end;

//==============================================================================

Function Ptr_WriteAnsiString_LE(var Dest: Pointer; const Str: AnsiString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_LE(WorkPtr,Length(Str),True);
If Length(Str) > 0 then
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PAnsiChar(Str)^,Length(Str) * SizeOf(AnsiChar),True));
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiString_LE(Dest: Pointer; const Str: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteAnsiString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiString_BE(var Dest: Pointer; const Str: AnsiString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_BE(WorkPtr,Length(Str),True);
If Length(Str) > 0 then
  Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PAnsiChar(Str)^,Length(Str) * SizeOf(AnsiChar),True));
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiString_BE(Dest: Pointer; const Str: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteAnsiString_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiString(var Dest: Pointer; const Str: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteAnsiString_BE(Dest,Str,Advance)
else
  Result := Ptr_WriteAnsiString_LE(Dest,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiString(Dest: Pointer; const Str: AnsiString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteAnsiString_BE(Ptr,Str,False)
else
  Result := Ptr_WriteAnsiString_LE(Ptr,Str,False);
end;

//==============================================================================

Function Ptr_WriteUTF8String_LE(var Dest: Pointer; const Str: UTF8String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_LE(WorkPtr,Length(Str),True);
If Length(Str) > 0 then
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PUTF8Char(Str)^,Length(Str) * SizeOf(UTF8Char),True));
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
Function Ptr_WriteUTF8String_LE(Dest: Pointer; const Str: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUTF8String_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8String_BE(var Dest: Pointer; const Str: UTF8String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_BE(WorkPtr,Length(Str),True);
If Length(Str) > 0 then
  Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PUTF8Char(Str)^,Length(Str) * SizeOf(UTF8Char),True));
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
Function Ptr_WriteUTF8String_BE(Dest: Pointer; const Str: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUTF8String_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8String(var Dest: Pointer; const Str: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUTF8String_BE(Dest,Str,Advance)
else
  Result := Ptr_WriteUTF8String_LE(Dest,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
Function Ptr_WriteUTF8String(Dest: Pointer; const Str: UTF8String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUTF8String_BE(Ptr,Str,False)
else
  Result := Ptr_WriteUTF8String_LE(Ptr,Str,False);
end;

//==============================================================================

Function Ptr_WriteWideString_LE(var Dest: Pointer; const Str: WideString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_LE(WorkPtr,Length(Str),True);
If Length(Str) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Str)),Length(Str)));
{$ELSE}
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PWideChar(Str)^,Length(Str) * SizeOf(WideChar),True));
{$ENDIF}
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideString_LE(Dest: Pointer; const Str: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteWideString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteWideString_BE(var Dest: Pointer; const Str: WideString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_BE(WorkPtr,Length(Str),True);
If Length(Str) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PWideChar(Str)^,Length(Str) * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Str)),Length(Str)));
{$ENDIF}
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideString_BE(Dest: Pointer; const Str: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteWideString_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteWideString(var Dest: Pointer; const Str: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteWideString_BE(Dest,Str,Advance)
else
  Result := Ptr_WriteWideString_LE(Dest,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideString(Dest: Pointer; const Str: WideString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteWideString_BE(Ptr,Str,False)
else
  Result := Ptr_WriteWideString_LE(Ptr,Str,False);
end;

//==============================================================================

Function Ptr_WriteUnicodeString_LE(var Dest: Pointer; const Str: UnicodeString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_LE(WorkPtr,Length(Str),True);
If Length(Str) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Str)),Length(Str)));
{$ELSE}
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PUnicodeChar(Str)^,Length(Str) * SizeOf(UnicodeChar),True));
{$ENDIF}
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeString_LE(Dest: Pointer; const Str: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUnicodeString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeString_BE(var Dest: Pointer; const Str: UnicodeString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_BE(WorkPtr,Length(Str),True);
If Length(Str) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PUnicodeChar(Str)^,Length(Str) * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Str)),Length(Str)));
{$ENDIF}
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeString_BE(Dest: Pointer; const Str: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUnicodeString_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeString(var Dest: Pointer; const Str: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUnicodeString_BE(Dest,Str,Advance)
else
  Result := Ptr_WriteUnicodeString_LE(Dest,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeString(Dest: Pointer; const Str: UnicodeString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUnicodeString_BE(Ptr,Str,False)
else
  Result := Ptr_WriteUnicodeString_LE(Ptr,Str,False);
end;

//==============================================================================

Function Ptr_WriteUCS4String_LE(var Dest: Pointer; const Str: UCS4String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
  TrueLen:  TStrSize;
begin
WorkPtr := Dest;
If Length(Str) > 0 then
  begin
    If Str[High(Str)] = 0 then
      TrueLen := Pred(Length(Str))
    else
      TrueLen := Length(Str);
    If TrueLen > 0 then
      begin
        Result := Ptr_WriteInt32_LE(WorkPtr,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Ptr_WriteUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Str[Low(Str)])),TrueLen));
      {$ELSE}
        Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,Addr(Str[Low(Str)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ENDIF}
      end
    else Result := Ptr_WriteInt32_LE(WorkPtr,0,True);
  end
else Result := Ptr_WriteInt32_LE(WorkPtr,0,True);
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4String_LE(Dest: Pointer; const Str: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUCS4String_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUCS4String_BE(var Dest: Pointer; const Str: UCS4String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
  TrueLen:  TStrSize;
begin
WorkPtr := Dest;
If Length(Str) > 0 then
  begin
    If Str[High(Str)] = 0 then
      TrueLen := Pred(Length(Str))
    else
      TrueLen := Length(Str);
    If TrueLen > 0 then
      begin
        Result := Ptr_WriteInt32_BE(WorkPtr,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,Addr(Str[Low(Str)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ELSE}
        Inc(Result,Ptr_WriteUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Str[Low(Str)])),TrueLen));
      {$ENDIF}
      end
    else Result := Ptr_WriteInt32_BE(WorkPtr,0,True);
  end
else Result := Ptr_WriteInt32_BE(WorkPtr,0,True);
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4String_BE(Dest: Pointer; const Str: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUCS4String_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUCS4String(var Dest: Pointer; const Str: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUCS4String_BE(Dest,Str,Advance)
else
  Result := Ptr_WriteUCS4String_LE(Dest,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4String(Dest: Pointer; const Str: UCS4String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUCS4String_BE(Ptr,Str,False)
else
  Result := Ptr_WriteUCS4String_LE(Ptr,Str,False);
end;

//==============================================================================

Function Ptr_WriteString_LE(var Dest: Pointer; const Str: String; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteUTF8String_LE(Dest,StrToUTF8(Str),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteString_LE(Dest: Pointer; const Str: String): TMemSize;
begin
Result := Ptr_WriteUTF8String_LE(Dest,StrToUTF8(Str));
end;

//------------------------------------------------------------------------------

Function Ptr_WriteString_BE(var Dest: Pointer; const Str: String; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteUTF8String_BE(Dest,StrToUTF8(Str),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteString_BE(Dest: Pointer; const Str: String): TMemSize;
begin
Result := Ptr_WriteUTF8String_BE(Dest,StrToUTF8(Str));
end;

//------------------------------------------------------------------------------

Function Ptr_WriteString(var Dest: Pointer; const Str: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteUTF8String(Dest,StrToUTF8(Str),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteString(Dest: Pointer; const Str: String; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteUTF8String(Dest,StrToUTF8(Str),Endian);
end;

{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function _Ptr_WriteBuffer(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Move(Buffer,Dest^,Size);
Result := Size;
AdvancePointer(Advance,Dest,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBuffer_LE(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBuffer(Dest,Buffer,Size,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
Function Ptr_WriteBuffer_LE(Dest: Pointer; const Buffer; Size: TMemSize): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteBuffer(Ptr,Buffer,Size,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBuffer_BE(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBuffer(Dest,Buffer,Size,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
Function Ptr_WriteBuffer_BE(Dest: Pointer; const Buffer; Size: TMemSize): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteBuffer(Ptr,Buffer,Size,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBuffer(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteBuffer_BE(Dest,Buffer,Size,Advance)
else
  Result := Ptr_WriteBuffer_LE(Dest,Buffer,Size,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
Function Ptr_WriteBuffer(Dest: Pointer; const Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteBuffer_BE(Ptr,Buffer,Size,False)
else
  Result := Ptr_WriteBuffer_LE(Ptr,Buffer,Size,False);
end;

//==============================================================================

Function _Ptr_WriteBytes(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize;
var
  WorkPtr:  PUInt8;
  i:        Integer;
begin
If ByteOpenArrayIsPacked then
  begin
    // write whole array
    If Length(Value) > 0 then
      Move(Value[Low(Value)],Dest^,Length(Value));
    Result := TMemSize(Length(Value));
    AdvancePointer(Advance,Dest,Result);
  end
else 
  begin
  {
    write byte-by-byte

    As we are writing into memory, there is no point in buffering the
    individual bytes.
  }
    WorkPtr := Dest;
    For i := Low(Value) to High(Value) do
      begin
        WorkPtr^ := Value[i];
        Inc(WorkPtr);
      end;
    Result := TMemSize(Length(Value));      
    If Advance then
      Dest := WorkPtr;
  end;
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBytes_LE(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBytes(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
Function Ptr_WriteBytes_LE(Dest: Pointer; const Value: array of UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteBytes(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBytes_BE(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBytes(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
Function Ptr_WriteBytes_BE(Dest: Pointer; const Value: array of UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteBytes(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteBytes(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteBytes_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteBytes_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteBytes(Dest: Pointer; const Value: array of UInt8; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteBytes_BE(Ptr,Value,False)
else
  Result := Ptr_WriteBytes_LE(Ptr,Value,False);
end;

//==============================================================================

Function _Ptr_FillBytes(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize;
begin
FillChar(Dest^,Count,Value);
Result := Count;
AdvancePointer(Advance,Dest,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_FillBytes_LE(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_FillBytes(Dest,Count,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_FillBytes_LE(Dest: Pointer; Count: TMemSize; Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_FillBytes(Ptr,Count,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_FillBytes_BE(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_FillBytes(Dest,Count,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_FillBytes_BE(Dest: Pointer; Count: TMemSize; Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_FillBytes(Ptr,Count,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_FillBytes(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_FillBytes_BE(Dest,Count,Value,Advance)
else
  Result := Ptr_FillBytes_LE(Dest,Count,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_FillBytes(Dest: Pointer; Count: TMemSize; Value: UInt8; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_FillBytes_BE(Ptr,Count,Value,False)
else
  Result := Ptr_FillBytes_LE(Ptr,Count,Value,False);
end;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Ptr_WriteVariant_LE(var Dest: Pointer; const Value: Variant; Advance: Boolean): TMemSize;
var
  Ptr:        Pointer;
  Dimensions: Integer;
  i:          Integer;       
  Indices:    array of Integer;

  Function WriteVarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:  Integer;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          Inc(Result,Ptr_WriteVariant_LE(Ptr,VarArrayGet(Value,Indices),True))
        else
          Inc(Result,WriteVarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
Ptr := Dest;
If VarIsArray(Value) then
  begin
    // array type
    Result := Ptr_WriteUInt8_LE(Ptr,VarTypeToInt(VarType(Value)),True);
    Dimensions := VarArrayDimCount(Value);
    // write number of dimensions
    Inc(Result,Ptr_WriteInt32_LE(Ptr,Dimensions,True));
    // write indices bounds (pairs of integers - low and high bound) for each dimension
    For i := 1 to Dimensions do
      begin
        Inc(Result,Ptr_WriteInt32_LE(Ptr,VarArrayLowBound(Value,i),True));
        Inc(Result,Ptr_WriteInt32_LE(Ptr,VarArrayHighBound(Value,i),True));
      end;
  {
    Now (recursively :P) write data if any exist.

    Btw. I know about VarArrayLock/VarArrayUnlock and pointer directly to the
    data, but I want it to be fool proof, therefore using VarArrayGet instead
    of direct access,
  }
    If Dimensions > 0 then
      begin
        SetLength(Indices,Dimensions);
        Inc(Result,WriteVarArrayDimension(1));
      end;
  end
else
  begin
    // simple type
    If VarType(Value) <> (varVariant or varByRef) then
      begin
        Result := Ptr_WriteUInt8_LE(Ptr,VarTypeToInt(VarType(Value)),True);
        case VarType(Value) and varTypeMask of
          varBoolean:   Inc(Result,Ptr_WriteBool_LE(Ptr,VarToBool(Value),True));
          varShortInt:  Inc(Result,Ptr_WriteInt8_LE(Ptr,Value,True));
          varSmallint:  Inc(Result,Ptr_WriteInt16_LE(Ptr,Value,True));
          varInteger:   Inc(Result,Ptr_WriteInt32_LE(Ptr,Value,True));
          varInt64:     Inc(Result,Ptr_WriteInt64_LE(Ptr,VarToInt64(Value),True));
          varByte:      Inc(Result,Ptr_WriteUInt8_LE(Ptr,Value,True));
          varWord:      Inc(Result,Ptr_WriteUInt16_LE(Ptr,Value,True));
          varLongWord:  Inc(Result,Ptr_WriteUInt32_LE(Ptr,Value,True));
        {$IF Declared(varUInt64)}
          varUInt64:    Inc(Result,Ptr_WriteUInt64_LE(Ptr,VarToUInt64(Value),True));
        {$IFEND}
          varSingle:    Inc(Result,Ptr_WriteFloat32_LE(Ptr,Value,True));
          varDouble:    Inc(Result,Ptr_WriteFloat64_LE(Ptr,Value,True));
          varCurrency:  Inc(Result,Ptr_WriteCurrency_LE(Ptr,Value,True));
          varDate:      Inc(Result,Ptr_WriteDateTime_LE(Ptr,Value,True));
          varOleStr:    Inc(Result,Ptr_WriteWideString_LE(Ptr,Value,True));
          varString:    Inc(Result,Ptr_WriteAnsiString_LE(Ptr,AnsiString(Value),True));
        {$IF Declared(varUString)}
          varUString:   Inc(Result,Ptr_WriteUnicodeString_LE(Ptr,Value,True));
        {$IFEND}
        else
          raise EBSUnsupportedVarType.CreateFmt('Ptr_WriteVariant_LE: Cannot write variant of this type (%d).',[VarType(Value) and varTypeMask]);
        end;
      end
    else Result := Ptr_WriteVariant_LE(Ptr,Variant(FindVarData(Value)^),True);
  end;
If Advance then
  Dest := Ptr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteVariant_LE(Dest: Pointer; const Value: Variant): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteVariant_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteVariant_BE(var Dest: Pointer; const Value: Variant; Advance: Boolean): TMemSize;
var
  Ptr:        Pointer;
  Dimensions: Integer;
  i:          Integer;       
  Indices:    array of Integer;

  Function WriteVarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:  Integer;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          Inc(Result,Ptr_WriteVariant_BE(Ptr,VarArrayGet(Value,Indices),True))
        else
          Inc(Result,WriteVarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
Ptr := Dest;
If VarIsArray(Value) then
  begin
    Result := Ptr_WriteUInt8_BE(Ptr,VarTypeToInt(VarType(Value)),True);
    Dimensions := VarArrayDimCount(Value);
    Inc(Result,Ptr_WriteInt32_BE(Ptr,Dimensions,True));
    For i := 1 to Dimensions do
      begin
        Inc(Result,Ptr_WriteInt32_BE(Ptr,VarArrayLowBound(Value,i),True));
        Inc(Result,Ptr_WriteInt32_BE(Ptr,VarArrayHighBound(Value,i),True));
      end;
    If Dimensions > 0 then
      begin
        SetLength(Indices,Dimensions);
        Inc(Result,WriteVarArrayDimension(1));
      end;
  end
else
  begin
    If VarType(Value) <> (varVariant or varByRef) then
      begin
        Result := Ptr_WriteUInt8_BE(Ptr,VarTypeToInt(VarType(Value)),True);
        case VarType(Value) and varTypeMask of
          varBoolean:   Inc(Result,Ptr_WriteBool_BE(Ptr,VarToBool(Value),True));
          varShortInt:  Inc(Result,Ptr_WriteInt8_BE(Ptr,Value,True));
          varSmallint:  Inc(Result,Ptr_WriteInt16_BE(Ptr,Value,True));
          varInteger:   Inc(Result,Ptr_WriteInt32_BE(Ptr,Value,True));
          varInt64:     Inc(Result,Ptr_WriteInt64_BE(Ptr,VarToInt64(Value),True));
          varByte:      Inc(Result,Ptr_WriteUInt8_BE(Ptr,Value,True));
          varWord:      Inc(Result,Ptr_WriteUInt16_BE(Ptr,Value,True));
          varLongWord:  Inc(Result,Ptr_WriteUInt32_BE(Ptr,Value,True));
        {$IF Declared(varUInt64)}
          varUInt64:    Inc(Result,Ptr_WriteUInt64_BE(Ptr,VarToUInt64(Value),True));
        {$IFEND}
          varSingle:    Inc(Result,Ptr_WriteFloat32_BE(Ptr,Value,True));
          varDouble:    Inc(Result,Ptr_WriteFloat64_BE(Ptr,Value,True));
          varCurrency:  Inc(Result,Ptr_WriteCurrency_BE(Ptr,Value,True));
          varDate:      Inc(Result,Ptr_WriteDateTime_BE(Ptr,Value,True));
          varOleStr:    Inc(Result,Ptr_WriteWideString_BE(Ptr,Value,True));
          varString:    Inc(Result,Ptr_WriteAnsiString_BE(Ptr,AnsiString(Value),True));
        {$IF Declared(varUString)}
          varUString:   Inc(Result,Ptr_WriteUnicodeString_BE(Ptr,Value,True));
        {$IFEND}
        else
          raise EBSUnsupportedVarType.CreateFmt('Ptr_WriteVariant_BE: Cannot write variant of this type (%d).',[VarType(Value) and varTypeMask]);
        end;
      end
    else Result := Ptr_WriteVariant_BE(Ptr,Variant(FindVarData(Value)^),True);
  end;
If Advance then
  Dest := Ptr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteVariant_BE(Dest: Pointer; const Value: Variant): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteVariant_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteVariant(var Dest: Pointer; const Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteVariant_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteVariant_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteVariant(Dest: Pointer; const Value: Variant; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteVariant_BE(Ptr,Value,False)
else
  Result := Ptr_WriteVariant_LE(Ptr,Value,False);
end;


{===============================================================================
--------------------------------------------------------------------------------
                                 Memory reading
--------------------------------------------------------------------------------
===============================================================================}

Function _Ptr_ReadBool(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize;
begin
Value := NumToBool(UInt8(Src^));
Result := SizeOf(UInt8);
AdvancePointer(Advance,Src,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadBool_LE(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadBool(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadBool_LE(Src: Pointer; out Value: ByteBool): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadBool(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadBool_BE(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadBool(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadBool_BE(Src: Pointer; out Value: ByteBool): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadBool(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadBool(var Src: Pointer; out Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadBool_BE(Src,Value,Advance)
else
  Result := Ptr_ReadBool_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
       
Function Ptr_ReadBool(Src: Pointer; out Value: ByteBool; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadBool_BE(Ptr,Value,False)
else
  Result := Ptr_ReadBool_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetBool_LE(var Src: Pointer; Advance: Boolean): ByteBool;
begin
_Ptr_ReadBool(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetBool_LE(Src: Pointer): ByteBool;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadBool(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetBool_BE(var Src: Pointer; Advance: Boolean): ByteBool;
begin
_Ptr_ReadBool(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetBool_BE(Src: Pointer): ByteBool;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadBool(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetBool(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): ByteBool;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetBool_BE(Src,Advance)
else
  Result := Ptr_GetBool_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
      
Function Ptr_GetBool(Src: Pointer; Endian: TEndian = endDefault): ByteBool;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetBool_BE(Ptr,False)
else
  Result := Ptr_GetBool_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadBoolean_LE(var Src: Pointer; out Value: Boolean; Advance: Boolean): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Ptr_ReadBool_LE(Src,TempBool,Advance);
Value := TempBool;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadBoolean_LE(Src: Pointer; out Value: Boolean): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Ptr_ReadBool_LE(Src,TempBool);
Value := TempBool;
end;

//------------------------------------------------------------------------------

Function Ptr_ReadBoolean_BE(var Src: Pointer; out Value: Boolean; Advance: Boolean): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Ptr_ReadBool_BE(Src,TempBool,Advance);
Value := TempBool;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadBoolean_BE(Src: Pointer; out Value: Boolean): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Ptr_ReadBool_BE(Src,TempBool);
Value := TempBool;
end;

//------------------------------------------------------------------------------

Function Ptr_ReadBoolean(var Src: Pointer; out Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Ptr_ReadBool(Src,TempBool,Advance,Endian);
Value := TempBool;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadBoolean(Src: Pointer; out Value: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Ptr_ReadBool(Src,TempBool,Endian);
Value := TempBool;
end;

//------------------------------------------------------------------------------

Function Ptr_GetBoolean_LE(var Src: Pointer; Advance: Boolean): Boolean;
begin
Result := Ptr_GetBool_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetBoolean_LE(Src: Pointer): Boolean;
begin
Result := Ptr_GetBool_LE(Src);
end;

//------------------------------------------------------------------------------

Function Ptr_GetBoolean_BE(var Src: Pointer; Advance: Boolean): Boolean;
begin
Result := Ptr_GetBool_BE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetBoolean_BE(Src: Pointer): Boolean;
begin
Result := Ptr_GetBool_BE(Src);
end;

//------------------------------------------------------------------------------

Function Ptr_GetBoolean(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Boolean;
begin
Result := Ptr_GetBool(Src,Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetBoolean(Src: Pointer; Endian: TEndian = endDefault): Boolean;
begin
Result := Ptr_GetBool(Src,Endian);
end;

{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function _Ptr_ReadInt8(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize;
begin
Value := Int8(Src^);
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt8_LE(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadInt8(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
Function Ptr_ReadInt8_LE(Src: Pointer; out Value: Int8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadInt8(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt8_BE(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadInt8(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
Function Ptr_ReadInt8_BE(Src: Pointer; out Value: Int8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadInt8(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt8(var Src: Pointer; out Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadInt8_BE(Src,Value,Advance)
else
  Result := Ptr_ReadInt8_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
Function Ptr_ReadInt8(Src: Pointer; out Value: Int8; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadInt8_BE(Ptr,Value,False)
else
  Result := Ptr_ReadInt8_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt8_LE(var Src: Pointer; Advance: Boolean): Int8;
begin
_Ptr_ReadInt8(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt8_LE(Src: Pointer): Int8;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadInt8(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt8_BE(var Src: Pointer; Advance: Boolean): Int8;
begin
_Ptr_ReadInt8(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt8_BE(Src: Pointer): Int8;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadInt8(Ptr,Result,False);
end; 

//------------------------------------------------------------------------------

Function Ptr_GetInt8(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int8;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetInt8_BE(Src,Advance)
else
  Result := Ptr_GetInt8_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt8(Src: Pointer; Endian: TEndian = endDefault): Int8;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetInt8_BE(Ptr,False)
else
  Result := Ptr_GetInt8_LE(Ptr,False);
end;

//==============================================================================

Function _Ptr_ReadUInt8(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize;
begin
Value := UInt8(Src^);
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt8_LE(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadUInt8(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt8_LE(Src: Pointer; out Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadUInt8(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt8_BE(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadUInt8(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt8_BE(Src: Pointer; out Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadUInt8(Ptr,Value,False);
end;

//------------------------------------------------------------------------------ 

Function Ptr_ReadUInt8(var Src: Pointer; out Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUInt8_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUInt8_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt8(Src: Pointer; out Value: UInt8; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUInt8_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUInt8_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt8_LE(var Src: Pointer; Advance: Boolean): UInt8;
begin
_Ptr_ReadUInt8(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt8_LE(Src: Pointer): UInt8;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadUInt8(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt8_BE(var Src: Pointer; Advance: Boolean): UInt8;
begin
_Ptr_ReadUInt8(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt8_BE(Src: Pointer): UInt8;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadUInt8(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt8(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt8;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUInt8_BE(Src,Advance)
else
  Result := Ptr_GetUInt8_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
               
Function Ptr_GetUInt8(Src: Pointer; Endian: TEndian = endDefault): UInt8;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUInt8_BE(Ptr,False)
else
  Result := Ptr_GetUInt8_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadInt16_LE(var Src: Pointer; out Value: Int16; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int16(SwapEndian(UInt16(Src^)));
{$ELSE}
Value := Int16(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt16_LE(Src: Pointer; out Value: Int16): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadInt16_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt16_BE(var Src: Pointer; out Value: Int16; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int16(Src^);
{$ELSE}
Value := Int16(SwapEndian(UInt16(Src^)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt16_BE(Src: Pointer; out Value: Int16): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadInt16_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt16(var Src: Pointer; out Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadInt16_BE(Src,Value,Advance)
else
  Result := Ptr_ReadInt16_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt16(Src: Pointer; out Value: Int16; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadInt16_BE(Ptr,Value,False)
else
  Result := Ptr_ReadInt16_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt16_LE(var Src: Pointer; Advance: Boolean): Int16;
begin
Ptr_ReadInt16_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt16_LE(Src: Pointer): Int16;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadInt16_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt16_BE(var Src: Pointer; Advance: Boolean): Int16;
begin
Ptr_ReadInt16_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt16_BE(Src: Pointer): Int16;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadInt16_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt16(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int16;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetInt16_BE(Src,Advance)
else
  Result := Ptr_GetInt16_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt16(Src: Pointer; Endian: TEndian = endDefault): Int16;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetInt16_BE(Ptr,False)
else
  Result := Ptr_GetInt16_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadUInt16_LE(var Src: Pointer; out Value: UInt16; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(UInt16(Src^));
{$ELSE}
Value := UInt16(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt16_LE(Src: Pointer; out Value: UInt16): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUInt16_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt16_BE(var Src: Pointer; out Value: UInt16; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UInt16(Src^);
{$ELSE}
Value := SwapEndian(UInt16(Src^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt16_BE(Src: Pointer; out Value: UInt16): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUInt16_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt16(var Src: Pointer; out Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUInt16_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUInt16_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt16(Src: Pointer; out Value: UInt16; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUInt16_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUInt16_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt16_LE(var Src: Pointer; Advance: Boolean): UInt16;
begin
Ptr_ReadUInt16_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt16_LE(Src: Pointer): UInt16;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUInt16_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt16_BE(var Src: Pointer; Advance: Boolean): UInt16;
begin
Ptr_ReadUInt16_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt16_BE(Src: Pointer): UInt16;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUInt16_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt16(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt16;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUInt16_BE(Src,Advance)
else
  Result := Ptr_GetUInt16_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt16(Src: Pointer; Endian: TEndian = endDefault): UInt16;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUInt16_BE(Ptr,False)
else
  Result := Ptr_GetUInt16_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadInt32_LE(var Src: Pointer; out Value: Int32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int32(SwapEndian(UInt32(Src^)));
{$ELSE}
Value := Int32(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt32_LE(Src: Pointer; out Value: Int32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadInt32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt32_BE(var Src: Pointer; out Value: Int32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int32(Src^);
{$ELSE}
Value := Int32(SwapEndian(UInt32(Src^)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt32_BE(Src: Pointer; out Value: Int32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadInt32_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt32(var Src: Pointer; out Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadInt32_BE(Src,Value,Advance)
else
  Result := Ptr_ReadInt32_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt32(Src: Pointer; out Value: Int32; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadInt32_BE(Ptr,Value,False)
else
  Result := Ptr_ReadInt32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt32_LE(var Src: Pointer; Advance: Boolean): Int32;
begin
Ptr_ReadInt32_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt32_LE(Src: Pointer): Int32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadInt32_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt32_BE(var Src: Pointer; Advance: Boolean): Int32;
begin
Ptr_ReadInt32_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt32_BE(Src: Pointer): Int32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadInt32_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetInt32_BE(Src,Advance)
else
  Result := Ptr_GetInt32_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt32(Src: Pointer; Endian: TEndian = endDefault): Int32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetInt32_BE(Ptr,False)
else
  Result := Ptr_GetInt32_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadUInt32_LE(var Src: Pointer; out Value: UInt32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(UInt32(Src^));
{$ELSE}
Value := UInt32(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt32_LE(Src: Pointer; out Value: UInt32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUInt32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt32_BE(var Src: Pointer; out Value: UInt32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UInt32(Src^);
{$ELSE}
Value := SwapEndian(UInt32(Src^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt32_BE(Src: Pointer; out Value: UInt32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUInt32_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt32(var Src: Pointer; out Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUInt32_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUInt32_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt32(Src: Pointer; out Value: UInt32; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUInt32_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUInt32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt32_LE(var Src: Pointer; Advance: Boolean): UInt32;
begin
Ptr_ReadUInt32_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt32_LE(Src: Pointer): UInt32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUInt32_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt32_BE(var Src: Pointer; Advance: Boolean): UInt32;
begin
Ptr_ReadUInt32_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt32_BE(Src: Pointer): UInt32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUInt32_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUInt32_BE(Src,Advance)
else
  Result := Ptr_GetUInt32_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt32(Src: Pointer; Endian: TEndian = endDefault): UInt32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUInt32_BE(Ptr,False)
else
  Result := Ptr_GetUInt32_LE(Ptr,False);
end;
 
//==============================================================================

Function Ptr_ReadInt64_LE(var Src: Pointer; out Value: Int64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int64(SwapEndian(UInt64(Src^)));
{$ELSE}
Value := Int64(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt64_LE(Src: Pointer; out Value: Int64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadInt64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt64_BE(var Src: Pointer; out Value: Int64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int64(Src^);
{$ELSE}
Value := Int64(SwapEndian(UInt64(Src^)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt64_BE(Src: Pointer; out Value: Int64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadInt64_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadInt64(var Src: Pointer; out Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadInt64_BE(Src,Value,Advance)
else
  Result := Ptr_ReadInt64_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadInt64(Src: Pointer; out Value: Int64; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadInt64_BE(Ptr,Value,False)
else
  Result := Ptr_ReadInt64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt64_LE(var Src: Pointer; Advance: Boolean): Int64;
begin
Ptr_ReadInt64_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt64_LE(Src: Pointer): Int64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadInt64_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt64_BE(var Src: Pointer; Advance: Boolean): Int64;
begin
Ptr_ReadInt64_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt64_BE(Src: Pointer): Int64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadInt64_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetInt64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetInt64_BE(Src,Advance)
else
  Result := Ptr_GetInt64_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetInt64(Src: Pointer; Endian: TEndian = endDefault): Int64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetInt64_BE(Ptr,False)
else
  Result := Ptr_GetInt64_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadUInt64_LE(var Src: Pointer; out Value: UInt64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(UInt64(Src^));
{$ELSE}
Value := UInt64(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt64_LE(Src: Pointer; out Value: UInt64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUInt64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt64_BE(var Src: Pointer; out Value: UInt64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UInt64(Src^);
{$ELSE}
Value := SwapEndian(UInt64(Src^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt64_BE(Src: Pointer; out Value: UInt64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUInt64_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUInt64(var Src: Pointer; out Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUInt64_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUInt64_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUInt64(Src: Pointer; out Value: UInt64; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUInt64_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUInt64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt64_LE(var Src: Pointer; Advance: Boolean): UInt64;
begin
Ptr_ReadUInt64_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt64_LE(Src: Pointer): UInt64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUInt64_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt64_BE(var Src: Pointer; Advance: Boolean): UInt64;
begin
Ptr_ReadUInt64_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt64_BE(Src: Pointer): UInt64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUInt64_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUInt64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUInt64_BE(Src,Advance)
else
  Result := Ptr_GetUInt64_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUInt64(Src: Pointer; Endian: TEndian = endDefault): UInt64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUInt64_BE(Ptr,False)
else
  Result := Ptr_GetUInt64_LE(Ptr,False);
end;

{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Ptr_ReadFloat32_LE(var Src: Pointer; out Value: Float32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt32(Addr(Value)^) := SwapEndian(UInt32(Src^));
{$ELSE}
Value := Float32(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat32_LE(Src: Pointer; out Value: Float32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadFloat32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadFloat32_BE(var Src: Pointer; out Value: Float32; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Float32(Src^);
{$ELSE}
UInt32(Addr(Value)^) := SwapEndian(UInt32(Src^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat32_BE(Src: Pointer; out Value: Float32): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadFloat32_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadFloat32(var Src: Pointer; out Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadFloat32_BE(Src,Value,Advance)
else
  Result := Ptr_ReadFloat32_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat32(Src: Pointer; out Value: Float32; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadFloat32_BE(Ptr,Value,False)
else
  Result := Ptr_ReadFloat32_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat32_LE(var Src: Pointer; Advance: Boolean): Float32;
begin
Ptr_ReadFloat32_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetFloat32_LE(Src: Pointer): Float32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadFloat32_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat32_BE(var Src: Pointer; Advance: Boolean): Float32;
begin
Ptr_ReadFloat32_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetFloat32_BE(Src: Pointer): Float32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadFloat32_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetFloat32_BE(Src,Advance)
else
  Result := Ptr_GetFloat32_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                 
Function Ptr_GetFloat32(Src: Pointer; Endian: TEndian = endDefault): Float32;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetFloat32_BE(Ptr,False)
else
  Result := Ptr_GetFloat32_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadFloat64_LE(var Src: Pointer; out Value: Float64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
UInt64(Addr(Value)^) := SwapEndian(UInt64(Src^));
{$ELSE}
Value := Float64(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat64_LE(Src: Pointer; out Value: Float64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadFloat64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadFloat64_BE(var Src: Pointer; out Value: Float64; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Float64(Src^);
{$ELSE}
UInt64(Addr(Value)^) := SwapEndian(UInt64(Src^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat64_BE(Src: Pointer; out Value: Float64): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadFloat64_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadFloat64(var Src: Pointer; out Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadFloat64_BE(Src,Value,Advance)
else
  Result := Ptr_ReadFloat64_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat64(Src: Pointer; out Value: Float64; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadFloat64_BE(Ptr,Value,False)
else
  Result := Ptr_ReadFloat64_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat64_LE(var Src: Pointer; Advance: Boolean): Float64;
begin
Ptr_ReadFloat64_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetFloat64_LE(Src: Pointer): Float64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadFloat64_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat64_BE(var Src: Pointer; Advance: Boolean): Float64;
begin
Ptr_ReadFloat64_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetFloat64_BE(Src: Pointer): Float64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadFloat64_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetFloat64_BE(Src,Advance)
else
  Result := Ptr_GetFloat64_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                 
Function Ptr_GetFloat64(Src: Pointer; Endian: TEndian = endDefault): Float64;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetFloat64_BE(Ptr,False)
else
  Result := Ptr_GetFloat64_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadFloat80_LE(var Src: Pointer; out Value: Float80; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
TFloat80Overlay(Addr(Value)^) := SwapEndian(TFloat80Overlay(Src^));
{$ELSE}
Value := Float80(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat80_LE(Src: Pointer; out Value: Float80): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadFloat80_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadFloat80_BE(var Src: Pointer; out Value: Float80; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Float80(Src^);
{$ELSE}
TFloat80Overlay(Addr(Value)^) := SwapEndian(TFloat80Overlay(Src^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat80_BE(Src: Pointer; out Value: Float80): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadFloat80_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadFloat80(var Src: Pointer; out Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadFloat80_BE(Src,Value,Advance)
else
  Result := Ptr_ReadFloat80_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadFloat80(Src: Pointer; out Value: Float80; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadFloat80_BE(Ptr,Value,False)
else
  Result := Ptr_ReadFloat80_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat80_LE(var Src: Pointer; Advance: Boolean): Float80;
begin
Ptr_ReadFloat80_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetFloat80_LE(Src: Pointer): Float80;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadFloat80_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat80_BE(var Src: Pointer; Advance: Boolean): Float80;
begin
Ptr_ReadFloat80_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetFloat80_BE(Src: Pointer): Float80;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadFloat80_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetFloat80(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float80;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetFloat80_BE(Src,Advance)
else
  Result := Ptr_GetFloat80_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                 
Function Ptr_GetFloat80(Src: Pointer; Endian: TEndian = endDefault): Float80;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetFloat80_BE(Ptr,False)
else
  Result := Ptr_GetFloat80_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadDateTime_LE(var Src: Pointer; out Value: TDateTime; Advance: Boolean): TMemSize;
begin
Result := Ptr_ReadFloat64_LE(Src,Float64(Value),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadDateTime_LE(Src: Pointer; out Value: TDateTime): TMemSize;
begin
Result := Ptr_ReadFloat64_LE(Src,Float64(Value));
end;

//------------------------------------------------------------------------------

Function Ptr_ReadDateTime_BE(var Src: Pointer; out Value: TDateTime; Advance: Boolean): TMemSize;
begin
Result := Ptr_ReadFloat64_BE(Src,Float64(Value),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadDateTime_BE(Src: Pointer; out Value: TDateTime): TMemSize;
begin
Result := Ptr_ReadFloat64_BE(Src,Float64(Value));
end;

//------------------------------------------------------------------------------

Function Ptr_ReadDateTime(var Src: Pointer; out Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_ReadFloat64(Src,Float64(Value),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadDateTime(Src: Pointer; out Value: TDateTime; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_ReadFloat64(Src,Float64(Value),Endian);
end;

//------------------------------------------------------------------------------

Function Ptr_GetDateTime_LE(var Src: Pointer; Advance: Boolean): TDateTime;
begin
Result := Ptr_GetFloat64_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetDateTime_LE(Src: Pointer): TDateTime;
begin
Result := Ptr_GetFloat64_LE(Src);
end;

//------------------------------------------------------------------------------

Function Ptr_GetDateTime_BE(var Src: Pointer; Advance: Boolean): TDateTime;
begin
Result := Ptr_GetFloat64_BE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetDateTime_BE(Src: Pointer): TDateTime;
begin
Result := Ptr_GetFloat64_BE(Src);
end;

//------------------------------------------------------------------------------

Function Ptr_GetDateTime(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): TDateTime;
begin
Result := Ptr_GetFloat64(Src,Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetDateTime(Src: Pointer; Endian: TEndian = endDefault): TDateTime;
begin
Result := Ptr_GetFloat64(Src,Endian);
end;

//==============================================================================

Function Ptr_ReadCurrency_LE(var Src: Pointer; out Value: Currency; Advance: Boolean): TMemSize;
{$IFDEF ENDIAN_BIG}
var
  Temp: UInt64 absolute Value;
{$ENDIF}
begin
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(UInt64(Src^));
{$ELSE}
Value := Currency(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadCurrency_LE(Src: Pointer; out Value: Currency): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadCurrency_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadCurrency_BE(var Src: Pointer; out Value: Currency; Advance: Boolean): TMemSize;
{$IFNDEF ENDIAN_BIG}
var
  Temp: UInt64 absolute Value;
{$ENDIF}
begin
{$IFDEF ENDIAN_BIG}
Value := Currency(Src^);
{$ELSE}
Temp := SwapEndian(UInt64(Src^));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadCurrency_BE(Src: Pointer; out Value: Currency): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadCurrency_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadCurrency(var Src: Pointer; out Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadCurrency_BE(Src,Value,Advance)
else
  Result := Ptr_ReadCurrency_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadCurrency(Src: Pointer; out Value: Currency; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadCurrency_BE(Ptr,Value,False)
else
  Result := Ptr_ReadCurrency_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetCurrency_LE(var Src: Pointer; Advance: Boolean): Currency;
begin
Ptr_ReadCurrency_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetCurrency_LE(Src: Pointer): Currency;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadCurrency_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetCurrency_BE(var Src: Pointer; Advance: Boolean): Currency;
begin
Ptr_ReadCurrency_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetCurrency_BE(Src: Pointer): Currency;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadCurrency_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetCurrency(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Currency;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetCurrency_BE(Src,Advance)
else
  Result := Ptr_GetCurrency_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetCurrency(Src: Pointer; Endian: TEndian = endDefault): Currency;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetCurrency_BE(Ptr,False)
else
  Result := Ptr_GetCurrency_LE(Ptr,False);
end;

{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function _Ptr_ReadAnsiChar(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Value := AnsiChar(Src^);
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadAnsiChar_LE(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadAnsiChar(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiChar_LE(Src: Pointer; out Value: AnsiChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadAnsiChar(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadAnsiChar_BE(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadAnsiChar(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiChar_BE(Src: Pointer; out Value: AnsiChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadAnsiChar(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadAnsiChar(var Src: Pointer; out Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadAnsiChar_BE(Src,Value,Advance)
else
  Result := Ptr_ReadAnsiChar_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiChar(Src: Pointer; out Value: AnsiChar; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadAnsiChar_BE(Ptr,Value,False)
else
  Result := Ptr_ReadAnsiChar_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetAnsiChar_LE(var Src: Pointer; Advance: Boolean): AnsiChar;
begin
_Ptr_ReadAnsiChar(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetAnsiChar_LE(Src: Pointer): AnsiChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadAnsiChar(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetAnsiChar_BE(var Src: Pointer; Advance: Boolean): AnsiChar;
begin
_Ptr_ReadAnsiChar(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetAnsiChar_BE(Src: Pointer): AnsiChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadAnsiChar(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetAnsiChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): AnsiChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetAnsiChar_BE(Src,Advance)
else
  Result := Ptr_GetAnsiChar_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetAnsiChar(Src: Pointer; Endian: TEndian = endDefault): AnsiChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetAnsiChar_BE(Ptr,False)
else
  Result := Ptr_GetAnsiChar_LE(Ptr,False);
end;

//==============================================================================

Function _Ptr_ReadUTF8Char(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Value := UTF8Char(Src^);
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8Char_LE(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadUTF8Char(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8Char_LE(Src: Pointer; out Value: UTF8Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadUTF8Char(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8Char_BE(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadUTF8Char(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8Char_BE(Src: Pointer; out Value: UTF8Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadUTF8Char(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8Char(var Src: Pointer; out Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUTF8Char_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUTF8Char_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8Char(Src: Pointer; out Value: UTF8Char; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUTF8Char_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUTF8Char_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUTF8Char_LE(var Src: Pointer; Advance: Boolean): UTF8Char;
begin
_Ptr_ReadUTF8Char(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUTF8Char_LE(Src: Pointer): UTF8Char;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadUTF8Char(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUTF8Char_BE(var Src: Pointer; Advance: Boolean): UTF8Char;
begin
_Ptr_ReadUTF8Char(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUTF8Char_BE(Src: Pointer): UTF8Char;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadUTF8Char(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUTF8Char(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UTF8Char;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUTF8Char_BE(Src,Advance)
else
  Result := Ptr_GetUTF8Char_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUTF8Char(Src: Pointer; Endian: TEndian = endDefault): UTF8Char;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUTF8Char_BE(Ptr,False)
else
  Result := Ptr_GetUTF8Char_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadWideChar_LE(var Src: Pointer; out Value: WideChar; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := WideChar(SwapEndian(UInt16(Src^)));
{$ELSE}
Value := WideChar(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadWideChar_LE(Src: Pointer; out Value: WideChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadWideChar_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadWideChar_BE(var Src: Pointer; out Value: WideChar; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := WideChar(Src^);
{$ELSE}
Value := WideChar(SwapEndian(UInt16(Src^)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadWideChar_BE(Src: Pointer; out Value: WideChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadWideChar_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadWideChar(var Src: Pointer; out Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadWideChar_BE(Src,Value,Advance)
else
  Result := Ptr_ReadWideChar_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                
Function Ptr_ReadWideChar(Src: Pointer; out Value: WideChar; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadWideChar_BE(Ptr,Value,False)
else
  Result := Ptr_ReadWideChar_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetWideChar_LE(var Src: Pointer; Advance: Boolean): WideChar;
begin
Ptr_ReadWideChar_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetWideChar_LE(Src: Pointer): WideChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadWideChar_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetWideChar_BE(var Src: Pointer; Advance: Boolean): WideChar;
begin
Ptr_ReadWideChar_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetWideChar_BE(Src: Pointer): WideChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadWideChar_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetWideChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): WideChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetWideChar_BE(Src,Advance)
else
  Result := Ptr_GetWideChar_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetWideChar(Src: Pointer; Endian: TEndian = endDefault): WideChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetWideChar_BE(Ptr,False)
else
  Result := Ptr_GetWideChar_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadUnicodeChar_LE(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UnicodeChar(SwapEndian(UInt16(Src^)));
{$ELSE}
Value := UnicodeChar(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUnicodeChar_LE(Src: Pointer; out Value: UnicodeChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUnicodeChar_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeChar_BE(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UnicodeChar(Src^);
{$ELSE}
Value := UnicodeChar(SwapEndian(UInt16(Src^)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUnicodeChar_BE(Src: Pointer; out Value: UnicodeChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUnicodeChar_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeChar(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUnicodeChar_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUnicodeChar_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                
Function Ptr_ReadUnicodeChar(Src: Pointer; out Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUnicodeChar_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUnicodeChar_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUnicodeChar_LE(var Src: Pointer; Advance: Boolean): UnicodeChar;
begin
Ptr_ReadUnicodeChar_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUnicodeChar_LE(Src: Pointer): UnicodeChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUnicodeChar_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUnicodeChar_BE(var Src: Pointer; Advance: Boolean): UnicodeChar;
begin
Ptr_ReadUnicodeChar_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUnicodeChar_BE(Src: Pointer): UnicodeChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUnicodeChar_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUnicodeChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UnicodeChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUnicodeChar_BE(Src,Advance)
else
  Result := Ptr_GetUnicodeChar_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUnicodeChar(Src: Pointer; Endian: TEndian = endDefault): UnicodeChar;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUnicodeChar_BE(Ptr,False)
else
  Result := Ptr_GetUnicodeChar_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadUCS4Char_LE(var Src: Pointer; out Value: UCS4Char; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UCS4Char(SwapEndian(UInt32(Src^)));
{$ELSE}
Value := UCS4Char(Src^);
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4Char_LE(Src: Pointer; out Value: UCS4Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUCS4Char_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUCS4Char_BE(var Src: Pointer; out Value: UCS4Char; Advance: Boolean): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UCS4Char(Src^);
{$ELSE}
Value := UCS4Char(SwapEndian(UInt32(Src^)));
{$ENDIF}
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4Char_BE(Src: Pointer; out Value: UCS4Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUCS4Char_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUCS4Char(var Src: Pointer; out Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUCS4Char_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUCS4Char_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4Char(Src: Pointer; out Value: UCS4Char; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUCS4Char_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUCS4Char_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUCS4Char_LE(var Src: Pointer; Advance: Boolean): UCS4Char;
begin
Ptr_ReadUCS4Char_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUCS4Char_LE(Src: Pointer): UCS4Char;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUCS4Char_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUCS4Char_BE(var Src: Pointer; Advance: Boolean): UCS4Char;
begin
Ptr_ReadUCS4Char_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUCS4Char_BE(Src: Pointer): UCS4Char;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUCS4Char_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUCS4Char(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UCS4Char;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUCS4Char_BE(Src,Advance)
else
  Result := Ptr_GetUCS4Char_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUCS4Char(Src: Pointer; Endian: TEndian = endDefault): UCS4Char;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUCS4Char_BE(Ptr,False)
else
  Result := Ptr_GetUCS4Char_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadChar_LE(var Src: Pointer; out Value: Char; Advance: Boolean): TMemSize;
var
  Temp: UInt16;
begin
Result := Ptr_ReadUInt16_LE(Src,Temp,Advance);
Value := Char(Temp);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadChar_LE(Src: Pointer; out Value: Char): TMemSize;
var
  Temp: UInt16;
begin
Result := Ptr_ReadUInt16_LE(Src,Temp);
Value := Char(Temp);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadChar_BE(var Src: Pointer; out Value: Char; Advance: Boolean): TMemSize;
var
  Temp: UInt16;
begin
Result := Ptr_ReadUInt16_BE(Src,Temp,Advance);
Value := Char(Temp);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadChar_BE(Src: Pointer; out Value: Char): TMemSize;
var
  Temp: UInt16;
begin
Result := Ptr_ReadUInt16_BE(Src,Temp);
Value := Char(Temp);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadChar(var Src: Pointer; out Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  Temp: UInt16;
begin
Result := Ptr_ReadUInt16(Src,Temp,Advance,Endian);
Value := Char(Temp);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadChar(Src: Pointer; out Value: Char; Endian: TEndian = endDefault): TMemSize;
var
  Temp: UInt16;
begin
Result := Ptr_ReadUInt16(Src,Temp,Endian);
Value := Char(Temp);
end;

//------------------------------------------------------------------------------

Function Ptr_GetChar_LE(var Src: Pointer; Advance: Boolean): Char;
begin
Result := Char(Ptr_GetUInt16_LE(Src,Advance));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetChar_LE(Src: Pointer): Char;
begin
Result := Char(Ptr_GetUInt16_LE(Src));
end;

//------------------------------------------------------------------------------

Function Ptr_GetChar_BE(var Src: Pointer; Advance: Boolean): Char;
begin
Result := Char(Ptr_GetUInt16_BE(Src,Advance));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetChar_BE(Src: Pointer): Char;
begin
Result := Char(Ptr_GetUInt16_BE(Src));
end;

//------------------------------------------------------------------------------

Function Ptr_GetChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Char;
begin
Result := Char(Ptr_GetUInt16(Src,Advance,Endian));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetChar(Src: Pointer; Endian: TEndian = endDefault): Char;
begin
Result := Char(Ptr_GetUInt16(Src,Endian));
end;

{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function _Ptr_ReadShortString(var Src: Pointer; out Str: ShortString; Advance: Boolean): TMemSize;
var
  StrLength:  UInt8;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadUInt8_LE(WorkPtr,StrLength,True);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,Addr(Str[1])^,StrLength,True));
If Advance then
  Src := WorkPtr;
end;

//------------------------------------------------------------------------------

Function Ptr_ReadShortString_LE(var Src: Pointer; out Str: ShortString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadShortString(Src,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadShortString_LE(Src: Pointer; out Str: ShortString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadShortString(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadShortString_BE(var Src: Pointer; out Str: ShortString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadShortString(Src,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadShortString_BE(Src: Pointer; out Str: ShortString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadShortString(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadShortString(var Src: Pointer; out Str: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadShortString_BE(Src,Str,Advance)
else
  Result := Ptr_ReadShortString_LE(Src,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadShortString(Src: Pointer; out Str: ShortString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadShortString_BE(Ptr,Str,False)
else
  Result := Ptr_ReadShortString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetShortString_LE(var Src: Pointer; Advance: Boolean): ShortString;
begin
_Ptr_ReadShortString(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetShortString_LE(Src: Pointer): ShortString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadShortString(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetShortString_BE(var Src: Pointer; Advance: Boolean): ShortString;
begin
_Ptr_ReadShortString(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetShortString_BE(Src: Pointer): ShortString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
_Ptr_ReadShortString(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetShortString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): ShortString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetShortString_BE(Src,Advance)
else
  Result := Ptr_GetShortString_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetShortString(Src: Pointer; Endian: TEndian = endDefault): ShortString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetShortString_BE(Ptr,False)
else
  Result := Ptr_GetShortString_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadAnsiString_LE(var Src: Pointer; out Str: AnsiString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PAnsiChar(Str)^,StrLength * SizeOf(AnsiChar),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiString_LE(Src: Pointer; out Str: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadAnsiString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadAnsiString_BE(var Src: Pointer; out Str: AnsiString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PAnsiChar(Str)^,StrLength * SizeOf(AnsiChar),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiString_BE(Src: Pointer; out Str: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadAnsiString_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadAnsiString(var Src: Pointer; out Str: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadAnsiString_BE(Src,Str,Advance)
else
  Result := Ptr_ReadAnsiString_LE(Src,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiString(Src: Pointer; out Str: AnsiString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadAnsiString_BE(Ptr,Str,False)
else
  Result := Ptr_ReadAnsiString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetAnsiString_LE(var Src: Pointer; Advance: Boolean): AnsiString;
begin
Ptr_ReadAnsiString_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetAnsiString_LE(Src: Pointer): AnsiString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadAnsiString_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetAnsiString_BE(var Src: Pointer; Advance: Boolean): AnsiString;
begin
Ptr_ReadAnsiString_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetAnsiString_BE(Src: Pointer): AnsiString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadAnsiString_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetAnsiString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): AnsiString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetAnsiString_BE(Src,Advance)
else
  Result := Ptr_GetAnsiString_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetAnsiString(Src: Pointer; Endian: TEndian = endDefault): AnsiString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetAnsiString_BE(Ptr,False)
else
  Result := Ptr_GetAnsiString_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadUTF8String_LE(var Src: Pointer; out Str: UTF8String; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PUTF8Char(Str)^,StrLength * SizeOf(UTF8Char),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8String_LE(Src: Pointer; out Str: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUTF8String_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8String_BE(var Src: Pointer; out Str: UTF8String; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PUTF8Char(Str)^,StrLength * SizeOf(UTF8Char),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8String_BE(Src: Pointer; out Str: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUTF8String_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8String(var Src: Pointer; out Str: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUTF8String_BE(Src,Str,Advance)
else
  Result := Ptr_ReadUTF8String_LE(Src,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8String(Src: Pointer; out Str: UTF8String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUTF8String_BE(Ptr,Str,False)
else
  Result := Ptr_ReadUTF8String_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUTF8String_LE(var Src: Pointer; Advance: Boolean): UTF8String;
begin
Ptr_ReadUTF8String_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUTF8String_LE(Src: Pointer): UTF8String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUTF8String_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUTF8String_BE(var Src: Pointer; Advance: Boolean): UTF8String;
begin
Ptr_ReadUTF8String_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUTF8String_BE(Src: Pointer): UTF8String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUTF8String_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUTF8String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UTF8String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUTF8String_BE(Src,Advance)
else
  Result := Ptr_GetUTF8String_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUTF8String(Src: Pointer; Endian: TEndian = endDefault): UTF8String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUTF8String_BE(Ptr,False)
else
  Result := Ptr_GetUTF8String_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadWideString_LE(var Src: Pointer; out Str: WideString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Str)),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PWideChar(Str)^,StrLength * SizeOf(WideChar),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadWideString_LE(Src: Pointer; out Str: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadWideString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadWideString_BE(var Src: Pointer; out Str: WideString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PWideChar(Str)^,StrLength * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Str)),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadWideString_BE(Src: Pointer; out Str: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadWideString_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadWideString(var Src: Pointer; out Str: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadWideString_BE(Src,Str,Advance)
else
  Result := Ptr_ReadWideString_LE(Src,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadWideString(Src: Pointer; out Str: WideString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadWideString_BE(Ptr,Str,False)
else
  Result := Ptr_ReadWideString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetWideString_LE(var Src: Pointer; Advance: Boolean): WideString;
begin
Ptr_ReadWideString_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetWideString_LE(Src: Pointer): WideString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadWideString_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetWideString_BE(var Src: Pointer; Advance: Boolean): WideString;
begin
Ptr_ReadWideString_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetWideString_BE(Src: Pointer): WideString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadWideString_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetWideString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): WideString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetWideString_BE(Src,Advance)
else
  Result := Ptr_GetWideString_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetWideString(Src: Pointer; Endian: TEndian = endDefault): WideString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetWideString_BE(Ptr,False)
else
  Result := Ptr_GetWideString_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadUnicodeString_LE(var Src: Pointer; out Str: UnicodeString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Str)),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PUnicodeChar(Str)^,StrLength * SizeOf(UnicodeChar),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUnicodeString_LE(Src: Pointer; out Str: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUnicodeString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeString_BE(var Src: Pointer; out Str: UnicodeString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PUnicodeChar(Str)^,StrLength * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Str)),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUnicodeString_BE(Src: Pointer; out Str: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUnicodeString_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeString(var Src: Pointer; out Str: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUnicodeString_BE(Src,Str,Advance)
else
  Result := Ptr_ReadUnicodeString_LE(Src,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUnicodeString(Src: Pointer; out Str: UnicodeString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUnicodeString_BE(Ptr,Str,False)
else
  Result := Ptr_ReadUnicodeString_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUnicodeString_LE(var Src: Pointer; Advance: Boolean): UnicodeString;
begin
Ptr_ReadUnicodeString_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUnicodeString_LE(Src: Pointer): UnicodeString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUnicodeString_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUnicodeString_BE(var Src: Pointer; Advance: Boolean): UnicodeString;
begin
Ptr_ReadUnicodeString_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUnicodeString_BE(Src: Pointer): UnicodeString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUnicodeString_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUnicodeString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUnicodeString_BE(Src,Advance)
else
  Result := Ptr_GetUnicodeString_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUnicodeString(Src: Pointer; Endian: TEndian = endDefault): UnicodeString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUnicodeString_BE(Ptr,False)
else
  Result := Ptr_GetUnicodeString_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadUCS4String_LE(var Src: Pointer; out Str: UCS4String; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength + 1);
Str[High(Str)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Str[Low(Str)])),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,Addr(Str[Low(Str)])^,StrLength * SizeOf(UCS4Char),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4String_LE(Src: Pointer; out Str: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUCS4String_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUCS4String_BE(var Src: Pointer; out Str: UCS4String; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength + 1);
Str[High(Str)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,Addr(Str[Low(Str)])^,StrLength * SizeOf(UCS4Char),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Str[Low(Str)])),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4String_BE(Src: Pointer; out Str: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUCS4String_BE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUCS4String(var Src: Pointer; out Str: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUCS4String_BE(Src,Str,Advance)
else
  Result := Ptr_ReadUCS4String_LE(Src,Str,Advance)
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4String(Src: Pointer; out Str: UCS4String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUCS4String_BE(Ptr,Str,False)
else
  Result := Ptr_ReadUCS4String_LE(Ptr,Str,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUCS4String_LE(var Src: Pointer; Advance: Boolean): UCS4String;
begin
Ptr_ReadUCS4String_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUCS4String_LE(Src: Pointer): UCS4String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUCS4String_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUCS4String_BE(var Src: Pointer; Advance: Boolean): UCS4String;
begin
Ptr_ReadUCS4String_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUCS4String_BE(Src: Pointer): UCS4String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadUCS4String_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetUCS4String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UCS4String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUCS4String_BE(Src,Advance)
else
  Result := Ptr_GetUCS4String_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetUCS4String(Src: Pointer; Endian: TEndian = endDefault): UCS4String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetUCS4String_BE(Ptr,False)
else
  Result := Ptr_GetUCS4String_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadString_LE(var Src: Pointer; out Str: String; Advance: Boolean): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String_LE(Src,TempStr,Advance);
Str := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadString_LE(Src: Pointer; out Str: String): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String_LE(Src,TempStr);
Str := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadString_BE(var Src: Pointer; out Str: String; Advance: Boolean): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String_BE(Src,TempStr,Advance);
Str := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadString_BE(Src: Pointer; out Str: String): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String_BE(Src,TempStr);
Str := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadString(var Src: Pointer; out Str: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String(Src,TempStr,Advance,Endian);
Str := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadString(Src: Pointer; out Str: String; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String(Src,TempStr,Endian);
Str := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Ptr_GetString_LE(var Src: Pointer; Advance: Boolean): String;
begin
Result := UTF8ToStr(Ptr_GetUTF8String_LE(Src,Advance));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetString_LE(Src: Pointer): String;
begin
Result := UTF8ToStr(Ptr_GetUTF8String_LE(Src));
end;

//------------------------------------------------------------------------------

Function Ptr_GetString_BE(var Src: Pointer; Advance: Boolean): String;
begin
Result := UTF8ToStr(Ptr_GetUTF8String_BE(Src,Advance));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetString_BE(Src: Pointer): String;
begin
Result := UTF8ToStr(Ptr_GetUTF8String_BE(Src));
end;

//------------------------------------------------------------------------------

Function Ptr_GetString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): String;
begin
Result := UTF8ToStr(Ptr_GetUTF8String(Src,Advance,Endian));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetString(Src: Pointer; Endian: TEndian = endDefault): String;
begin
Result := UTF8ToStr(Ptr_GetUTF8String(Src,Endian));
end;

{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function _Ptr_ReadBuffer(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
If Size > 0 then
  begin
    Move(Src^,Addr(Buffer)^,Size);
    Result := Size;
    AdvancePointer(Advance,Src,Result);
  end
else Result := 0;
end;

//------------------------------------------------------------------------------

Function Ptr_ReadBuffer_LE(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadBuffer(Src,Buffer,Size,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadBuffer_LE(Src: Pointer; out Buffer; Size: TMemSize): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadBuffer(Ptr,Buffer,Size,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadBuffer_BE(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadBuffer(Src,Buffer,Size,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadBuffer_BE(Src: Pointer; out Buffer; Size: TMemSize): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadBuffer(Ptr,Buffer,Size,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadBuffer(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadBuffer_BE(Src,Buffer,Size,Advance)
else
  Result := Ptr_ReadBuffer_LE(Src,Buffer,Size,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadBuffer(Src: Pointer; out Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadBuffer_BE(Ptr,Buffer,Size,False)
else
  Result := Ptr_ReadBuffer_LE(Ptr,Buffer,Size,False);
end;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Ptr_ReadVariant_LE(var Src: Pointer; out Value: Variant; Advance: Boolean): TMemSize;
var
  Ptr:            Pointer;
  VariantTypeInt: UInt8;
  Dimensions:     Integer;
  i:              Integer;
  IndicesBounds:  array of PtrInt;
  Indices:        array of Integer;
  WideStrTemp:    WideString;
  AnsiStrTemp:    AnsiString;
  UnicodeStrTemp: UnicodeString;

  Function ReadVarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:    Integer;
    TempVar:  Variant;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          begin
            Inc(Result,Ptr_ReadVariant_LE(Ptr,TempVar,True));
            VarArrayPut(Value,TempVar,Indices);
          end
        else Inc(Result,ReadVarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
Ptr := Src;
VariantTypeInt := Ptr_GetUInt8_LE(Ptr,True);
If VariantTypeInt and BS_VARTYPENUM_ARRAY <> 0 then
  begin
    // array
    Dimensions := Ptr_GetInt32_LE(Ptr,True);
    Result := 5 + (Dimensions * 8);
    If Dimensions > 0 then
      begin
        // read indices bounds
        SetLength(IndicesBounds,Dimensions * 2);
        For i := Low(IndicesBounds) to High(IndicesBounds) do
          IndicesBounds[i] := PtrInt(Ptr_GetInt32_LE(Ptr,True));
        // create the array
        Value := VarArrayCreate(IndicesBounds,IntToVarType(VariantTypeInt and BS_VARTYPENUM_TYPEMASK));
        // read individual dimensions/items
        SetLength(Indices,Dimensions);
        Inc(Result,ReadVarArrayDimension(1));
      end;
  end
else
  begin
    // simple type
    TVarData(Value).vType := IntToVarType(VariantTypeInt);
    case VariantTypeInt and BS_VARTYPENUM_TYPEMASK of
      BS_VARTYPENUM_BOOLEN:   begin
                                TVarData(Value).vBoolean := Ptr_GetBool_LE(Ptr,True);
                                // +1 is for the already read variable type byte
                                Result := StreamedSize_Bool + 1;
                              end;
      BS_VARTYPENUM_SHORTINT: Result := Ptr_ReadInt8_LE(Ptr,TVarData(Value).vShortInt,True) + 1;
      BS_VARTYPENUM_SMALLINT: Result := Ptr_ReadInt16_LE(Ptr,TVarData(Value).vSmallInt,True) + 1;
      BS_VARTYPENUM_INTEGER:  Result := Ptr_ReadInt32_LE(Ptr,TVarData(Value).vInteger,True) + 1;
      BS_VARTYPENUM_INT64:    Result := Ptr_ReadInt64_LE(Ptr,TVarData(Value).vInt64,True) + 1;
      BS_VARTYPENUM_BYTE:     Result := Ptr_ReadUInt8_LE(Ptr,TVarData(Value).vByte,True) + 1;
      BS_VARTYPENUM_WORD:     Result := Ptr_ReadUInt16_LE(Ptr,TVarData(Value).vWord,True) + 1;
      BS_VARTYPENUM_LONGWORD: Result := Ptr_ReadUInt32_LE(Ptr,TVarData(Value).vLongWord,True) + 1;
      BS_VARTYPENUM_UINT64: {$IF Declared(varUInt64)}
                              Result := Ptr_ReadUInt64_LE(Ptr,TVarData(Value).{$IFDEF FPC}vQWord{$ELSE}vUInt64{$ENDIF},True) + 1;
                            {$ELSE}
                              Result := Ptr_ReadUInt64_LE(Ptr,UInt64(TVarData(Value).vInt64),True) + 1;
                            {$IFEND}
      BS_VARTYPENUM_SINGLE:   Result := Ptr_ReadFloat32_LE(Ptr,TVarData(Value).vSingle,True) + 1;
      BS_VARTYPENUM_DOUBLE:   Result := Ptr_ReadFloat64_LE(Ptr,TVarData(Value).vDouble,True) + 1;
      BS_VARTYPENUM_CURRENCY: Result := Ptr_ReadCurrency_LE(Ptr,TVarData(Value).vCurrency,True) + 1;
      BS_VARTYPENUM_DATE:     Result := Ptr_ReadDateTime_LE(Ptr,TVarData(Value).vDate,True) + 1;
      BS_VARTYPENUM_OLESTR:   begin
                                Result := Ptr_ReadWideString_LE(Ptr,WideStrTemp,True) + 1;
                                Value := VarAsType(WideStrTemp,varOleStr);
                              end;
      BS_VARTYPENUM_STRING:   begin
                                Result := Ptr_ReadAnsiString_LE(Ptr,AnsiStrTemp,True) + 1;
                                Value := VarAsType(AnsiStrTemp,varString);
                              end;
      BS_VARTYPENUM_USTRING:  begin
                                Result := Ptr_ReadUnicodeString_LE(Ptr,UnicodeStrTemp,True) + 1;
                              {$IF Declared(varUString)}
                                Value := VarAsType(UnicodeStrTemp,varUString);
                              {$ELSE}
                                Value := VarAsType(UnicodeStrTemp,varOleStr);
                              {$IFEND}
                              end;
    else
      raise EBSUnsupportedVarType.CreateFmt('Ptr_ReadVariant_LE: Cannot read variant of this type number (%d).',[VariantTypeInt and BS_VARTYPENUM_TYPEMASK]);
    end;
  end;
If Advance then
  Src := Ptr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadVariant_LE(Src: Pointer; out Value: Variant): TMemSize; overload;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadVariant_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadVariant_BE(var Src: Pointer; out Value: Variant; Advance: Boolean): TMemSize;
var
  Ptr:            Pointer;
  VariantTypeInt: UInt8;
  Dimensions:     Integer;
  i:              Integer;
  IndicesBounds:  array of PtrInt;
  Indices:        array of Integer;
  WideStrTemp:    WideString;
  AnsiStrTemp:    AnsiString;
  UnicodeStrTemp: UnicodeString;

  Function ReadVarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:    Integer;
    TempVar:  Variant;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          begin
            Inc(Result,Ptr_ReadVariant_BE(Ptr,TempVar,True));
            VarArrayPut(Value,TempVar,Indices);
          end
        else Inc(Result,ReadVarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
Ptr := Src;
VariantTypeInt := Ptr_GetUInt8_BE(Ptr,True);
If VariantTypeInt and BS_VARTYPENUM_ARRAY <> 0 then
  begin
    // array
    Dimensions := Ptr_GetInt32_BE(Ptr,True);
    Result := 5 + (Dimensions * 8);
    If Dimensions > 0 then
      begin
        // read indices bounds
        SetLength(IndicesBounds,Dimensions * 2);
        For i := Low(IndicesBounds) to High(IndicesBounds) do
          IndicesBounds[i] := PtrInt(Ptr_GetInt32_BE(Ptr,True));
        // create the array
        Value := VarArrayCreate(IndicesBounds,IntToVarType(VariantTypeInt and BS_VARTYPENUM_TYPEMASK));
        // read individual dimensions/items
        SetLength(Indices,Dimensions);
        Inc(Result,ReadVarArrayDimension(1));
      end;
  end
else
  begin
    // simple type
    TVarData(Value).vType := IntToVarType(VariantTypeInt);
    case VariantTypeInt and BS_VARTYPENUM_TYPEMASK of
      BS_VARTYPENUM_BOOLEN:   begin
                                TVarData(Value).vBoolean := Ptr_GetBool_BE(Ptr,True);
                                // +1 is for the already read variable type byte
                                Result := StreamedSize_Bool + 1;
                              end;
      BS_VARTYPENUM_SHORTINT: Result := Ptr_ReadInt8_BE(Ptr,TVarData(Value).vShortInt,True) + 1;
      BS_VARTYPENUM_SMALLINT: Result := Ptr_ReadInt16_BE(Ptr,TVarData(Value).vSmallInt,True) + 1;
      BS_VARTYPENUM_INTEGER:  Result := Ptr_ReadInt32_BE(Ptr,TVarData(Value).vInteger,True) + 1;
      BS_VARTYPENUM_INT64:    Result := Ptr_ReadInt64_BE(Ptr,TVarData(Value).vInt64,True) + 1;
      BS_VARTYPENUM_BYTE:     Result := Ptr_ReadUInt8_BE(Ptr,TVarData(Value).vByte,True) + 1;
      BS_VARTYPENUM_WORD:     Result := Ptr_ReadUInt16_BE(Ptr,TVarData(Value).vWord,True) + 1;
      BS_VARTYPENUM_LONGWORD: Result := Ptr_ReadUInt32_BE(Ptr,TVarData(Value).vLongWord,True) + 1;
      BS_VARTYPENUM_UINT64: {$IF Declared(varUInt64)}
                              Result := Ptr_ReadUInt64_BE(Ptr,TVarData(Value).{$IFDEF FPC}vQWord{$ELSE}vUInt64{$ENDIF},True) + 1;
                            {$ELSE}
                              Result := Ptr_ReadUInt64_BE(Ptr,UInt64(TVarData(Value).vInt64),True) + 1;
                            {$IFEND}
      BS_VARTYPENUM_SINGLE:   Result := Ptr_ReadFloat32_BE(Ptr,TVarData(Value).vSingle,True) + 1;
      BS_VARTYPENUM_DOUBLE:   Result := Ptr_ReadFloat64_BE(Ptr,TVarData(Value).vDouble,True) + 1;
      BS_VARTYPENUM_CURRENCY: Result := Ptr_ReadCurrency_BE(Ptr,TVarData(Value).vCurrency,True) + 1;
      BS_VARTYPENUM_DATE:     Result := Ptr_ReadDateTime_BE(Ptr,TVarData(Value).vDate,True) + 1;
      BS_VARTYPENUM_OLESTR:   begin
                                Result := Ptr_ReadWideString_BE(Ptr,WideStrTemp,True) + 1;
                                Value := VarAsType(WideStrTemp,varOleStr);
                              end;
      BS_VARTYPENUM_STRING:   begin
                                Result := Ptr_ReadAnsiString_BE(Ptr,AnsiStrTemp,True) + 1;
                                Value := VarAsType(AnsiStrTemp,varString);
                              end;
      BS_VARTYPENUM_USTRING:  begin
                                Result := Ptr_ReadUnicodeString_BE(Ptr,UnicodeStrTemp,True) + 1;
                              {$IF Declared(varUString)}
                                Value := VarAsType(UnicodeStrTemp,varUString);
                              {$ELSE}
                                Value := VarAsType(UnicodeStrTemp,varOleStr);
                              {$IFEND}
                              end;
    else
      raise EBSUnsupportedVarType.CreateFmt('Ptr_ReadVariant_BE: Cannot read variant of this type number (%d).',[VariantTypeInt and BS_VARTYPENUM_TYPEMASK]);
    end;
  end;
If Advance then
  Src := Ptr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadVariant_BE(Src: Pointer; out Value: Variant): TMemSize; overload;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadVariant_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadVariant(var Src: Pointer; out Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadVariant_BE(Src,Value,Advance)
else
  Result := Ptr_ReadVariant_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadVariant(Src: Pointer; out Value: Variant; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadVariant_BE(Ptr,Value,False)
else
  Result := Ptr_ReadVariant_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetVariant_LE(var Src: Pointer; Advance: Boolean): Variant;
begin
Ptr_ReadVariant_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetVariant_LE(Src: Pointer): Variant;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadVariant_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetVariant_BE(var Src: Pointer; Advance: Boolean): Variant;
begin
Ptr_ReadVariant_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetVariant_BE(Src: Pointer): Variant;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadVariant_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetVariant(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Variant;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetVariant_BE(Src,Advance)
else
  Result := Ptr_GetVariant_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetVariant(Src: Pointer; Endian: TEndian = endDefault): Variant;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetVariant_BE(Ptr,False)
else
  Result := Ptr_GetVariant_LE(Ptr,False);
end;


{===============================================================================
--------------------------------------------------------------------------------
                                 Stream writing
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function _Stream_WriteBool(Stream: TStream; Value: ByteBool; Advance: Boolean): TMemSize;
var
  Temp: UInt8;
begin
Temp := BoolToNum(Value);
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Temp);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBool_LE(Stream: TStream; Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteBool(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBool_BE(Stream: TStream; Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteBool(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBool(Stream: TStream; Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteBool_BE(Stream,Value,Advance)
else
  Result := Stream_WriteBool_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteBool(Stream: TStream; Value: ByteBool; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteBool_BE(Stream,Value)
else
  Result := Stream_WriteBool_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteBoolean_LE(Stream: TStream; Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteBool_LE(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBoolean_BE(Stream: TStream; Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteBool_BE(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBoolean(Stream: TStream; Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteBool(Stream,Value,Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteBoolean(Stream: TStream; Value: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteBool(Stream,Value,Endian);
end;

{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function _Stream_WriteInt8(Stream: TStream; Value: Int8; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt8_LE(Stream: TStream; Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteInt8(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt8_BE(Stream: TStream; Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteInt8(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt8(Stream: TStream; Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteInt8_BE(Stream,Value,Advance)
else
  Result := Stream_WriteInt8_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteInt8(Stream: TStream; Value: Int8; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteInt8_BE(Stream,Value)
else
  Result := Stream_WriteInt8_LE(Stream,Value);
end;

//==============================================================================

Function _Stream_WriteUInt8(Stream: TStream; Value: UInt8; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt8_LE(Stream: TStream; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteUInt8(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt8_BE(Stream: TStream; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteUInt8(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt8(Stream: TStream; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUInt8_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUInt8_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUInt8(Stream: TStream; Value: UInt8; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUInt8_BE(Stream,Value)
else
  Result := Stream_WriteUInt8_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteInt16_LE(Stream: TStream; Value: Int16; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int16(SwapEndian(UInt16(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt16_BE(Stream: TStream; Value: Int16; Advance: Boolean = True): TMemSize;
begin
{$IFNDEF ENDIAN_BIG}
Value := Int16(SwapEndian(UInt16(Value))); 
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt16(Stream: TStream; Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteInt16_BE(Stream,Value,Advance)
else
  Result := Stream_WriteInt16_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteInt16(Stream: TStream; Value: Int16; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteInt16_BE(Stream,Value)
else
  Result := Stream_WriteInt16_LE(Stream,Value);
end;
 
//==============================================================================

Function Stream_WriteUInt16_LE(Stream: TStream; Value: UInt16; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt16_BE(Stream: TStream; Value: UInt16; Advance: Boolean = True): TMemSize;
begin
{$IFNDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt16(Stream: TStream; Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUInt16_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUInt16_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUInt16(Stream: TStream; Value: UInt16; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUInt16_BE(Stream,Value)
else
  Result := Stream_WriteUInt16_LE(Stream,Value);
end;
 
//==============================================================================

Function Stream_WriteInt32_LE(Stream: TStream; Value: Int32; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int32(SwapEndian(UInt32(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt32_BE(Stream: TStream; Value: Int32; Advance: Boolean = True): TMemSize;
begin
{$IFNDEF ENDIAN_BIG}
Value := Int32(SwapEndian(UInt32(Value))); 
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt32(Stream: TStream; Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteInt32_BE(Stream,Value,Advance)
else
  Result := Stream_WriteInt32_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteInt32(Stream: TStream; Value: Int32; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteInt32_BE(Stream,Value)
else
  Result := Stream_WriteInt32_LE(Stream,Value);
end;
 
//==============================================================================

Function Stream_WriteUInt32_LE(Stream: TStream; Value: UInt32; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt32_BE(Stream: TStream; Value: UInt32; Advance: Boolean = True): TMemSize;
begin
{$IFNDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt32(Stream: TStream; Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUInt32_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUInt32_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUInt32(Stream: TStream; Value: UInt32; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUInt32_BE(Stream,Value)
else
  Result := Stream_WriteUInt32_LE(Stream,Value);
end;
 
//==============================================================================

Function Stream_WriteInt64_LE(Stream: TStream; Value: Int64; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int64(SwapEndian(UInt64(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt64_BE(Stream: TStream; Value: Int64; Advance: Boolean = True): TMemSize;
begin
{$IFNDEF ENDIAN_BIG}
Value := Int64(SwapEndian(UInt64(Value))); 
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteInt64(Stream: TStream; Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteInt64_BE(Stream,Value,Advance)
else
  Result := Stream_WriteInt64_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteInt64(Stream: TStream; Value: Int64; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteInt64_BE(Stream,Value)
else
  Result := Stream_WriteInt64_LE(Stream,Value);
end;
 
//==============================================================================

Function Stream_WriteUInt64_LE(Stream: TStream; Value: UInt64; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt64_BE(Stream: TStream; Value: UInt64; Advance: Boolean = True): TMemSize;
begin
{$IFNDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUInt64(Stream: TStream; Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUInt64_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUInt64_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUInt64(Stream: TStream; Value: UInt64; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUInt64_BE(Stream,Value)
else
  Result := Stream_WriteUInt64_LE(Stream,Value);
end;
 
{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Stream_WriteFloat32_LE(Stream: TStream; Value: Float32; Advance: Boolean = True): TMemSize;
var
  Temp: UInt32 absolute Value;
begin
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteFloat32_BE(Stream: TStream; Value: Float32; Advance: Boolean = True): TMemSize;
var
  Temp: UInt32 absolute Value;
begin
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteFloat32(Stream: TStream; Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteFloat32_BE(Stream,Value,Advance)
else
  Result := Stream_WriteFloat32_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteFloat32(Stream: TStream; Value: Float32; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteFloat32_BE(Stream,Value)
else
  Result := Stream_WriteFloat32_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteFloat64_LE(Stream: TStream; Value: Float64; Advance: Boolean = True): TMemSize;
var
  Temp: UInt64 absolute Value;
begin
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteFloat64_BE(Stream: TStream; Value: Float64; Advance: Boolean = True): TMemSize;
var
  Temp: UInt64 absolute Value;
begin
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteFloat64(Stream: TStream; Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteFloat64_BE(Stream,Value,Advance)
else
  Result := Stream_WriteFloat64_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteFloat64(Stream: TStream; Value: Float64; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteFloat64_BE(Stream,Value)
else
  Result := Stream_WriteFloat64_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteFloat80_LE(Stream: TStream; Value: Float80; Advance: Boolean = True): TMemSize;
var
  Temp: TFloat80Overlay absolute Value;
begin
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteFloat80_BE(Stream: TStream; Value: Float80; Advance: Boolean = True): TMemSize;
var
  Temp: TFloat80Overlay absolute Value;
begin
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteFloat80(Stream: TStream; Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteFloat80_BE(Stream,Value,Advance)
else
  Result := Stream_WriteFloat80_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteFloat80(Stream: TStream; Value: Float80; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteFloat80_BE(Stream,Value)
else
  Result := Stream_WriteFloat80_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteDateTime_LE(Stream: TStream; Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteFloat64_LE(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteDateTime_BE(Stream: TStream; Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteFloat64_BE(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteDateTime(Stream: TStream; Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteFloat64(Stream,Value,Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteDateTime(Stream: TStream; Value: TDateTime; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteFloat64(Stream,Value,Endian);
end;

//==============================================================================

Function Stream_WriteCurrency_LE(Stream: TStream; Value: Currency; Advance: Boolean = True): TMemSize;
var
  Temp: UInt64 absolute Value;
begin
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteCurrency_BE(Stream: TStream; Value: Currency; Advance: Boolean = True): TMemSize;
var
  Temp: UInt64 absolute Value;
begin
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteCurrency(Stream: TStream; Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteCurrency_BE(Stream,Value,Advance)
else
  Result := Stream_WriteCurrency_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteCurrency(Stream: TStream; Value: Currency; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteCurrency_BE(Stream,Value)
else
  Result := Stream_WriteCurrency_LE(Stream,Value);
end;

{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function _Stream_WriteAnsiChar(Stream: TStream; Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteAnsiChar_LE(Stream: TStream; Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteAnsiChar(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteAnsiChar_BE(Stream: TStream; Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteAnsiChar(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteAnsiChar(Stream: TStream; Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteAnsiChar_BE(Stream,Value,Advance)
else
  Result := Stream_WriteAnsiChar_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteAnsiChar(Stream: TStream; Value: AnsiChar; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteAnsiChar_BE(Stream,Value)
else
  Result := Stream_WriteAnsiChar_LE(Stream,Value);
end;

//==============================================================================

Function _Stream_WriteUTF8Char(Stream: TStream; Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUTF8Char_LE(Stream: TStream; Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteUTF8Char(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUTF8Char_BE(Stream: TStream; Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteUTF8Char(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUTF8Char(Stream: TStream; Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUTF8Char_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUTF8Char_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUTF8Char(Stream: TStream; Value: UTF8Char; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUTF8Char_BE(Stream,Value)
else
  Result := Stream_WriteUTF8Char_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteWideChar_LE(Stream: TStream; Value: WideChar; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := WideChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteWideChar_BE(Stream: TStream; Value: WideChar; Advance: Boolean = True): TMemSize;
begin
{$IFNDEF ENDIAN_BIG}
Value := WideChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteWideChar(Stream: TStream; Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteWideChar_BE(Stream,Value,Advance)
else
  Result := Stream_WriteWideChar_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteWideChar(Stream: TStream; Value: WideChar; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteWideChar_BE(Stream,Value)
else
  Result := Stream_WriteWideChar_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteUnicodeChar_LE(Stream: TStream; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UnicodeChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUnicodeChar_BE(Stream: TStream; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
{$IFNDEF ENDIAN_BIG}
Value := UnicodeChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUnicodeChar(Stream: TStream; Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUnicodeChar_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUnicodeChar_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUnicodeChar(Stream: TStream; Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUnicodeChar_BE(Stream,Value)
else
  Result := Stream_WriteUnicodeChar_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteUCS4Char_LE(Stream: TStream; Value: UCS4Char; Advance: Boolean = True): TMemSize;
var
  Temp: UInt32 absolute Value;
begin
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUCS4Char_BE(Stream: TStream; Value: UCS4Char; Advance: Boolean = True): TMemSize;
var
  Temp: UInt32 absolute Value;
begin
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUCS4Char(Stream: TStream; Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUCS4Char_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUCS4Char_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUCS4Char(Stream: TStream; Value: UCS4Char; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUCS4Char_BE(Stream,Value)
else
  Result := Stream_WriteUCS4Char_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteChar_LE(Stream: TStream; Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteUInt16_LE(Stream,UInt16(Ord(Value)),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteChar_BE(Stream: TStream; Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteUInt16_BE(Stream,UInt16(Ord(Value)),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteChar(Stream: TStream; Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteUInt16(Stream,UInt16(Ord(Value)),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteChar(Stream: TStream; Value: Char; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteUInt16(Stream,UInt16(Ord(Value)),Endian);
end;

{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function _Stream_WriteShortString(Stream: TStream; const Str: ShortString; Advance: Boolean): TMemSize;
begin
Result := Stream_WriteUInt8_LE(Stream,UInt8(Length(Str)),True);
If Length(Str) > 0 then
  Inc(Result,Stream_WriteBuffer_LE(Stream,Addr(Str[1])^,Length(Str),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteShortString_LE(Stream: TStream; const Str: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteShortString(Stream,Str,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteShortString_BE(Stream: TStream; const Str: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteShortString(Stream,Str,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteShortString(Stream: TStream; const Str: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteShortString_BE(Stream,Str,Advance)
else
  Result := Stream_WriteShortString_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteShortString(Stream: TStream; const Str: ShortString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteShortString_BE(Stream,Str)
else
  Result := Stream_WriteShortString_LE(Stream,Str);
end;

//==============================================================================

Function Stream_WriteAnsiString_LE(Stream: TStream; const Str: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_LE(Stream,Length(Str),True);
If Length(Str) > 0 then
  Inc(Result,Stream_WriteBuffer_LE(Stream,PAnsiChar(Str)^,Length(Str) * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteAnsiString_BE(Stream: TStream; const Str: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_BE(Stream,Length(Str),True);
If Length(Str) > 0 then
  Inc(Result,Stream_WriteBuffer_BE(Stream,PAnsiChar(Str)^,Length(Str) * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteAnsiString(Stream: TStream; const Str: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteAnsiString_BE(Stream,Str,Advance)
else
  Result := Stream_WriteAnsiString_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteAnsiString(Stream: TStream; const Str: AnsiString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteAnsiString_BE(Stream,Str)
else
  Result := Stream_WriteAnsiString_LE(Stream,Str);
end;

//==============================================================================

Function Stream_WriteUTF8String_LE(Stream: TStream; const Str: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_LE(Stream,Length(Str),True);
If Length(Str) > 0 then
  Inc(Result,Stream_WriteBuffer_LE(Stream,PUTF8Char(Str)^,Length(Str) * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUTF8String_BE(Stream: TStream; const Str: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_BE(Stream,Length(Str),True);
If Length(Str) > 0 then
  Inc(Result,Stream_WriteBuffer_BE(Stream,PUTF8Char(Str)^,Length(Str) * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUTF8String(Stream: TStream; const Str: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUTF8String_BE(Stream,Str,Advance)
else
  Result := Stream_WriteUTF8String_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUTF8String(Stream: TStream; const Str: UTF8String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUTF8String_BE(Stream,Str)
else
  Result := Stream_WriteUTF8String_LE(Stream,Str);
end;

//==============================================================================

Function Stream_WriteWideString_LE(Stream: TStream; const Str: WideString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_LE(Stream,Length(Str),True);
If Length(Str) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Str)),Length(Str)));
{$ELSE}
  Inc(Result,Stream_WriteBuffer_LE(Stream,PWideChar(Str)^,Length(Str) * SizeOf(WideChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteWideString_BE(Stream: TStream; const Str: WideString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_BE(Stream,Length(Str),True);
If Length(Str) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_WriteBuffer_BE(Stream,PWideChar(Str)^,Length(Str) * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Str)),Length(Str)));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteWideString(Stream: TStream; const Str: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteWideString_BE(Stream,Str,Advance)
else
  Result := Stream_WriteWideString_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteWideString(Stream: TStream; const Str: WideString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteWideString_BE(Stream,Str)
else
  Result := Stream_WriteWideString_LE(Stream,Str);
end;

//==============================================================================

Function Stream_WriteUnicodeString_LE(Stream: TStream; const Str: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_LE(Stream,Length(Str),True);
If Length(Str) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Str)),Length(Str)));
{$ELSE}
  Inc(Result,Stream_WriteBuffer_LE(Stream,PUnicodeChar(Str)^,Length(Str) * SizeOf(UnicodeChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUnicodeString_BE(Stream: TStream; const Str: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_BE(Stream,Length(Str),True);
If Length(Str) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_WriteBuffer_BE(Stream,PUnicodeChar(Str)^,Length(Str) * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Str)),Length(Str)));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUnicodeString(Stream: TStream; const Str: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUnicodeString_BE(Stream,Str,Advance)
else
  Result := Stream_WriteUnicodeString_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUnicodeString(Stream: TStream; const Str: UnicodeString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUnicodeString_BE(Stream,Str)
else
  Result := Stream_WriteUnicodeString_LE(Stream,Str);
end;

//==============================================================================

Function Stream_WriteUCS4String_LE(Stream: TStream; const Str: UCS4String; Advance: Boolean = True): TMemSize;
var
  TrueLen:  TStrSize;
begin
If Length(Str) > 0 then
  begin
    If Str[High(Str)] = 0 then
      TrueLen := Pred(Length(Str))
    else
      TrueLen := Length(Str);
    If TrueLen > 0 then
      begin
        Result := Stream_WriteInt32_LE(Stream,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Stream_WriteUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Str[Low(Str)])),TrueLen));
      {$ELSE}
        Inc(Result,Stream_WriteBuffer_LE(Stream,Addr(Str[Low(Str)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ENDIF}
      end
    else Result := Stream_WriteInt32_LE(Stream,0,True);
  end
else Result := Stream_WriteInt32_LE(Stream,0,True);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUCS4String_BE(Stream: TStream; const Str: UCS4String; Advance: Boolean = True): TMemSize;
var
  TrueLen:  TStrSize;
begin
If Length(Str) > 0 then
  begin
    If Str[High(Str)] = 0 then
      TrueLen := Pred(Length(Str))
    else
      TrueLen := Length(Str);
    If TrueLen > 0 then
      begin
        Result := Stream_WriteInt32_BE(Stream,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Stream_WriteBuffer_BE(Stream,Addr(Str[Low(Str)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ELSE}
        Inc(Result,Stream_WriteUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Str[Low(Str)])),TrueLen));
      {$ENDIF}
      end
    else Result := Stream_WriteInt32_BE(Stream,0,True);
  end
else Result := Stream_WriteInt32_BE(Stream,0,True);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUCS4String(Stream: TStream; const Str: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUCS4String_BE(Stream,Str,Advance)
else
  Result := Stream_WriteUCS4String_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUCS4String(Stream: TStream; const Str: UCS4String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUCS4String_BE(Stream,Str)
else
  Result := Stream_WriteUCS4String_LE(Stream,Str);
end;

//==============================================================================

Function Stream_WriteString_LE(Stream: TStream; const Str: String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteUTF8String_LE(Stream,StrToUTF8(Str),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteString_BE(Stream: TStream; const Str: String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteUTF8String_BE(Stream,StrToUTF8(Str),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteString(Stream: TStream; const Str: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteUTF8String(Stream,StrToUTF8(Str),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteString(Stream: TStream; const Str: String; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteUTF8String(Stream,StrToUTF8(Str),Endian);
end;

{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function _Stream_WriteBuffer(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Buffer,Size);
Result := Size;
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBuffer_LE(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteBuffer(Stream,Buffer,Size,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBuffer_BE(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteBuffer(Stream,Buffer,Size,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBuffer(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteBuffer_BE(Stream,Buffer,Size,Advance)
else
  Result := Stream_WriteBuffer_LE(Stream,Buffer,Size,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteBuffer(Stream: TStream; const Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteBuffer_BE(Stream,Buffer,Size)
else
  Result := Stream_WriteBuffer_LE(Stream,Buffer,Size);
end;

//==============================================================================

Function _Stream_WriteBytes(Stream: TStream; const Value: array of UInt8; Advance: Boolean): TMemSize;
var
  Buffer:   packed array[0..1023] of UInt8; // buffer is on stack, keep it small
  Remain:   Integer;
  Offset:   Integer;
  CopyCnt:  Integer;
  i:        Integer;
begin
If not ByteOpenArrayIsPacked then
  begin
    Remain := Length(Value);
    Offset := 0;
    while Remain > 0 do
      begin
        CopyCnt := Min(Remain,SizeOf(Buffer));
        For i := 0 to Pred(CopyCnt) do
          Buffer[i] := Value[Offset + i];
        Stream.WriteBuffer(Addr(Buffer)^,CopyCnt);
        Dec(Remain,CopyCnt);
        Inc(Offset,CopyCnt);
      end;
  end
else Stream.WriteBuffer(Value[Low(Value)],Length(Value));
Result := TMemSize(Length(Value));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBytes_LE(Stream: TStream; const Value: array of UInt8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteBytes(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBytes_BE(Stream: TStream; const Value: array of UInt8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteBytes(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteBytes(Stream: TStream; const Value: array of UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteBytes_BE(Stream,Value,Advance)
else
  Result := Stream_WriteBytes_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteBytes(Stream: TStream; const Value: array of UInt8; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteBytes_BE(Stream,Value)
else
  Result := Stream_WriteBytes_LE(Stream,Value);
end;

//==============================================================================

Function _Stream_FillBytes(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize;
var
  Buffer:   packed array[0..1023] of UInt8;
  CopyCnt:  Integer;
begin
If Count >= SizeOf(Buffer) then
  FillChar(Addr(Buffer)^,SizeOf(Buffer),Value)
else
  FillChar(Addr(Buffer)^,Count,Value);
Result := 0;
while Count > 0 do
  begin
    CopyCnt := Min(Count,SizeOf(Buffer));
    Stream.WriteBuffer(Buffer,CopyCnt);
    Dec(Count,TMemSize(CopyCnt));
    Inc(Result,TMemSize(CopyCnt));
  end;
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_FillBytes_LE(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_FillBytes(Stream,Count,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_FillBytes_BE(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_FillBytes(Stream,Count,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_FillBytes(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_FillBytes_BE(Stream,Count,Value,Advance)
else
  Result := Stream_FillBytes_LE(Stream,Count,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_FillBytes(Stream: TStream; Count: TMemSize; Value: UInt8; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_FillBytes_BE(Stream,Count,Value)
else
  Result := Stream_FillBytes_LE(Stream,Count,Value);
end;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Stream_WriteVariant_LE(Stream: TStream; const Value: Variant; Advance: Boolean = True): TMemSize;
var
  Dimensions: Integer;
  i:          Integer;       
  Indices:    array of Integer;

  Function WriteVarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:  Integer;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          Inc(Result,Stream_WriteVariant_LE(Stream,VarArrayGet(Value,Indices),True))
        else
          Inc(Result,WriteVarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
If VarIsArray(Value) then
  begin
    Result := Stream_WriteUInt8_LE(Stream,VarTypeToInt(VarType(Value)),True);
    Dimensions := VarArrayDimCount(Value);
    Inc(Result,Stream_WriteInt32_LE(Stream,Dimensions,True));
    For i := 1 to Dimensions do
      begin
        Inc(Result,Stream_WriteInt32_LE(Stream,VarArrayLowBound(Value,i),True));
        Inc(Result,Stream_WriteInt32_LE(Stream,VarArrayHighBound(Value,i),True));
      end;
    If Dimensions > 0 then
      begin
        SetLength(Indices,Dimensions);
        Inc(Result,WriteVarArrayDimension(1));
      end;
  end
else
  begin
    If VarType(Value) <> (varVariant or varByRef) then
      begin
        Result := Stream_WriteUInt8_LE(Stream,VarTypeToInt(VarType(Value)),True);
        case VarType(Value) and varTypeMask of
          varBoolean:   Inc(Result,Stream_WriteBool_LE(Stream,VarToBool(Value),True));
          varShortInt:  Inc(Result,Stream_WriteInt8_LE(Stream,Value,True));
          varSmallint:  Inc(Result,Stream_WriteInt16_LE(Stream,Value,True));
          varInteger:   Inc(Result,Stream_WriteInt32_LE(Stream,Value,True));
          varInt64:     Inc(Result,Stream_WriteInt64_LE(Stream,VarToInt64(Value),True));
          varByte:      Inc(Result,Stream_WriteUInt8_LE(Stream,Value,True));
          varWord:      Inc(Result,Stream_WriteUInt16_LE(Stream,Value,True));
          varLongWord:  Inc(Result,Stream_WriteUInt32_LE(Stream,Value,True));
        {$IF Declared(varUInt64)}
          varUInt64:    Inc(Result,Stream_WriteUInt64_LE(Stream,VarToUInt64(Value),True));
        {$IFEND}
          varSingle:    Inc(Result,Stream_WriteFloat32_LE(Stream,Value,True));
          varDouble:    Inc(Result,Stream_WriteFloat64_LE(Stream,Value,True));
          varCurrency:  Inc(Result,Stream_WriteCurrency_LE(Stream,Value,True));
          varDate:      Inc(Result,Stream_WriteDateTime_LE(Stream,Value,True));
          varOleStr:    Inc(Result,Stream_WriteWideString_LE(Stream,Value,True));
          varString:    Inc(Result,Stream_WriteAnsiString_LE(Stream,AnsiString(Value),True));
        {$IF Declared(varUString)}
          varUString:   Inc(Result,Stream_WriteUnicodeString_LE(Stream,Value,True));
        {$IFEND}
        else
          raise EBSUnsupportedVarType.CreateFmt('Stream_WriteVariant_LE: Cannot write variant of this type (%d).',[VarType(Value) and varTypeMask]);
        end;
      end
    else Result := Stream_WriteVariant_LE(Stream,Variant(FindVarData(Value)^),True);
  end;
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteVariant_BE(Stream: TStream; const Value: Variant; Advance: Boolean = True): TMemSize;
var
  Dimensions: Integer;
  i:          Integer;       
  Indices:    array of Integer;

  Function WriteVarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:  Integer;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          Inc(Result,Stream_WriteVariant_BE(Stream,VarArrayGet(Value,Indices),True))
        else
          Inc(Result,WriteVarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
If VarIsArray(Value) then
  begin
    Result := Stream_WriteUInt8_BE(Stream,VarTypeToInt(VarType(Value)),True);
    Dimensions := VarArrayDimCount(Value);
    Inc(Result,Stream_WriteInt32_BE(Stream,Dimensions,True));
    For i := 1 to Dimensions do
      begin
        Inc(Result,Stream_WriteInt32_BE(Stream,VarArrayLowBound(Value,i),True));
        Inc(Result,Stream_WriteInt32_BE(Stream,VarArrayHighBound(Value,i),True));
      end;
    If Dimensions > 0 then
      begin
        SetLength(Indices,Dimensions);
        Inc(Result,WriteVarArrayDimension(1));
      end;
  end
else
  begin
    If VarType(Value) <> (varVariant or varByRef) then
      begin
        Result := Stream_WriteUInt8_BE(Stream,VarTypeToInt(VarType(Value)),True);
        case VarType(Value) and varTypeMask of
          varBoolean:   Inc(Result,Stream_WriteBool_BE(Stream,VarToBool(Value),True));
          varShortInt:  Inc(Result,Stream_WriteInt8_BE(Stream,Value,True));
          varSmallint:  Inc(Result,Stream_WriteInt16_BE(Stream,Value,True));
          varInteger:   Inc(Result,Stream_WriteInt32_BE(Stream,Value,True));
          varInt64:     Inc(Result,Stream_WriteInt64_BE(Stream,VarToInt64(Value),True));
          varByte:      Inc(Result,Stream_WriteUInt8_BE(Stream,Value,True));
          varWord:      Inc(Result,Stream_WriteUInt16_BE(Stream,Value,True));
          varLongWord:  Inc(Result,Stream_WriteUInt32_BE(Stream,Value,True));
        {$IF Declared(varUInt64)}
          varUInt64:    Inc(Result,Stream_WriteUInt64_BE(Stream,VarToUInt64(Value),True));
        {$IFEND}
          varSingle:    Inc(Result,Stream_WriteFloat32_BE(Stream,Value,True));
          varDouble:    Inc(Result,Stream_WriteFloat64_BE(Stream,Value,True));
          varCurrency:  Inc(Result,Stream_WriteCurrency_BE(Stream,Value,True));
          varDate:      Inc(Result,Stream_WriteDateTime_BE(Stream,Value,True));
          varOleStr:    Inc(Result,Stream_WriteWideString_BE(Stream,Value,True));
          varString:    Inc(Result,Stream_WriteAnsiString_BE(Stream,AnsiString(Value),True));
        {$IF Declared(varUString)}
          varUString:   Inc(Result,Stream_WriteUnicodeString_BE(Stream,Value,True));
        {$IFEND}
        else
          raise EBSUnsupportedVarType.CreateFmt('Stream_WriteVariant_BE: Cannot write variant of this type (%d).',[VarType(Value) and varTypeMask]);
        end;
      end
    else Result := Stream_WriteVariant_BE(Stream,Variant(FindVarData(Value)^),True);
  end;
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteVariant(Stream: TStream; const Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteVariant_BE(Stream,Value,Advance)
else
  Result := Stream_WriteVariant_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteVariant(Stream: TStream; const Value: Variant; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteVariant_BE(Stream,Value)
else
  Result := Stream_WriteVariant_LE(Stream,Value);
end;


{===============================================================================
--------------------------------------------------------------------------------
                                 Stream reading
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function _Stream_ReadBool(Stream: TStream; out Value: ByteBool; Advance: Boolean): TMemSize;
var
  Temp: UInt8;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
Result := SizeOf(Temp);
Value := NumToBool(Temp);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadBool_LE(Stream: TStream; out Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadBool(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadBool_BE(Stream: TStream; out Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadBool(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadBool(Stream: TStream; out Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadBool_BE(Stream,Value,Advance)
else
  Result := Stream_ReadBool_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadBool(Stream: TStream; out Value: ByteBool; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadBool_BE(Stream,Value)
else
  Result := Stream_ReadBool_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetBool_LE(Stream: TStream; Advance: Boolean = True): ByteBool;
begin
_Stream_ReadBool(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetBool_BE(Stream: TStream; Advance: Boolean = True): ByteBool;
begin
_Stream_ReadBool(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetBool(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): ByteBool;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetBool_BE(Stream,Advance)
else
  Result := Stream_GetBool_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetBool(Stream: TStream; Endian: TEndian = endDefault): ByteBool;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetBool_BE(Stream)
else
  Result := Stream_GetBool_LE(Stream);
end;

//==============================================================================

Function Stream_ReadBoolean_LE(Stream: TStream; out Value: Boolean; Advance: Boolean = True): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Stream_ReadBool_LE(Stream,TempBool,Advance);
Value := TempBool;
end;

//------------------------------------------------------------------------------

Function Stream_ReadBoolean_BE(Stream: TStream; out Value: Boolean; Advance: Boolean = True): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Stream_ReadBool_BE(Stream,TempBool,Advance);
Value := TempBool;
end;

//------------------------------------------------------------------------------

Function Stream_ReadBoolean(Stream: TStream; out Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Stream_ReadBool(Stream,TempBool,Advance,Endian);
Value := TempBool;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadBoolean(Stream: TStream; out Value: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Stream_ReadBool(Stream,TempBool,Endian);
Value := TempBool;
end;

//------------------------------------------------------------------------------

Function Stream_GetBoolean_LE(Stream: TStream; Advance: Boolean = True): Boolean;
begin
Result := Stream_GetBool_LE(Stream,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetBoolean_BE(Stream: TStream; Advance: Boolean = True): Boolean;
begin
Result := Stream_GetBool_BE(Stream,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetBoolean(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Boolean;
begin
Result := Stream_GetBool(Stream,Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetBoolean(Stream: TStream; Endian: TEndian = endDefault): Boolean;
begin
Result := Stream_GetBool(Stream,Endian);
end;

{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function _Stream_ReadInt8(Stream: TStream; out Value: Int8; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt8_LE(Stream: TStream; out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadInt8(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt8_BE(Stream: TStream; out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadInt8(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt8(Stream: TStream; out Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadInt8_BE(Stream,Value,Advance)
else
  Result := Stream_ReadInt8_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadInt8(Stream: TStream; out Value: Int8; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadInt8_BE(Stream,Value)
else
  Result := Stream_ReadInt8_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt8_LE(Stream: TStream; Advance: Boolean = True): Int8;
begin
_Stream_ReadInt8(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt8_BE(Stream: TStream; Advance: Boolean = True): Int8;
begin
_Stream_ReadInt8(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt8(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int8;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetInt8_BE(Stream,Advance)
else
  Result := Stream_GetInt8_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetInt8(Stream: TStream; Endian: TEndian = endDefault): Int8;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetInt8_BE(Stream)
else
  Result := Stream_GetInt8_LE(Stream);
end;

//==============================================================================

Function _Stream_ReadUInt8(Stream: TStream; out Value: UInt8; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt8_LE(Stream: TStream; out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadUInt8(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt8_BE(Stream: TStream; out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadUInt8(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt8(Stream: TStream; out Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUInt8_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUInt8_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUInt8(Stream: TStream; out Value: UInt8; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUInt8_BE(Stream,Value)
else
  Result := Stream_ReadUInt8_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt8_LE(Stream: TStream; Advance: Boolean = True): UInt8;
begin
_Stream_ReadUInt8(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt8_BE(Stream: TStream; Advance: Boolean = True): UInt8;
begin
_Stream_ReadUInt8(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt8(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt8;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUInt8_BE(Stream,Advance)
else
  Result := Stream_GetUInt8_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUInt8(Stream: TStream; Endian: TEndian = endDefault): UInt8;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUInt8_BE(Stream)
else
  Result := Stream_GetUInt8_LE(Stream);
end;

//==============================================================================

Function Stream_ReadInt16_LE(Stream: TStream; out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := Int16(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt16_BE(Stream: TStream; out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := Int16(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt16(Stream: TStream; out Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadInt16_BE(Stream,Value,Advance)
else
  Result := Stream_ReadInt16_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadInt16(Stream: TStream; out Value: Int16; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadInt16_BE(Stream,Value)
else
  Result := Stream_ReadInt16_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt16_LE(Stream: TStream; Advance: Boolean = True): Int16;
begin
Stream_ReadInt16_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt16_BE(Stream: TStream; Advance: Boolean = True): Int16;
begin
Stream_ReadInt16_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt16(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int16;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetInt16_BE(Stream,Advance)
else
  Result := Stream_GetInt16_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetInt16(Stream: TStream; Endian: TEndian = endDefault): Int16;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetInt16_BE(Stream)
else
  Result := Stream_GetInt16_LE(Stream);
end;

//==============================================================================

Function Stream_ReadUInt16_LE(Stream: TStream; out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt16_BE(Stream: TStream; out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt16(Stream: TStream; out Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUInt16_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUInt16_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUInt16(Stream: TStream; out Value: UInt16; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUInt16_BE(Stream,Value)
else
  Result := Stream_ReadUInt16_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt16_LE(Stream: TStream; Advance: Boolean = True): UInt16;
begin
Stream_ReadUInt16_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt16_BE(Stream: TStream; Advance: Boolean = True): UInt16;
begin
Stream_ReadUInt16_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt16(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt16;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUInt16_BE(Stream,Advance)
else
  Result := Stream_GetUInt16_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUInt16(Stream: TStream; Endian: TEndian = endDefault): UInt16;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUInt16_BE(Stream)
else
  Result := Stream_GetUInt16_LE(Stream);
end;

//==============================================================================

Function Stream_ReadInt32_LE(Stream: TStream; out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := Int32(SwapEndian(UInt32(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt32_BE(Stream: TStream; out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := Int32(SwapEndian(UInt32(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt32(Stream: TStream; out Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadInt32_BE(Stream,Value,Advance)
else
  Result := Stream_ReadInt32_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadInt32(Stream: TStream; out Value: Int32; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadInt32_BE(Stream,Value)
else
  Result := Stream_ReadInt32_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt32_LE(Stream: TStream; Advance: Boolean = True): Int32;
begin
Stream_ReadInt32_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt32_BE(Stream: TStream; Advance: Boolean = True): Int32;
begin
Stream_ReadInt32_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetInt32_BE(Stream,Advance)
else
  Result := Stream_GetInt32_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetInt32(Stream: TStream; Endian: TEndian = endDefault): Int32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetInt32_BE(Stream)
else
  Result := Stream_GetInt32_LE(Stream);
end;

//==============================================================================

Function Stream_ReadUInt32_LE(Stream: TStream; out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt32_BE(Stream: TStream; out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt32(Stream: TStream; out Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUInt32_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUInt32_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUInt32(Stream: TStream; out Value: UInt32; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUInt32_BE(Stream,Value)
else
  Result := Stream_ReadUInt32_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt32_LE(Stream: TStream; Advance: Boolean = True): UInt32;
begin
Stream_ReadUInt32_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt32_BE(Stream: TStream; Advance: Boolean = True): UInt32;
begin
Stream_ReadUInt32_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUInt32_BE(Stream,Advance)
else
  Result := Stream_GetUInt32_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUInt32(Stream: TStream; Endian: TEndian = endDefault): UInt32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUInt32_BE(Stream)
else
  Result := Stream_GetUInt32_LE(Stream);
end;

//==============================================================================

Function Stream_ReadInt64_LE(Stream: TStream; out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := Int64(SwapEndian(UInt64(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt64_BE(Stream: TStream; out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := Int64(SwapEndian(UInt64(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadInt64(Stream: TStream; out Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadInt64_BE(Stream,Value,Advance)
else
  Result := Stream_ReadInt64_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadInt64(Stream: TStream; out Value: Int64; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadInt64_BE(Stream,Value)
else
  Result := Stream_ReadInt64_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt64_LE(Stream: TStream; Advance: Boolean = True): Int64;
begin
Stream_ReadInt64_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt64_BE(Stream: TStream; Advance: Boolean = True): Int64;
begin
Stream_ReadInt64_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetInt64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetInt64_BE(Stream,Advance)
else
  Result := Stream_GetInt64_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetInt64(Stream: TStream; Endian: TEndian = endDefault): Int64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetInt64_BE(Stream)
else
  Result := Stream_GetInt64_LE(Stream);
end;

//==============================================================================

Function Stream_ReadUInt64_LE(Stream: TStream; out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt64_BE(Stream: TStream; out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUInt64(Stream: TStream; out Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUInt64_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUInt64_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUInt64(Stream: TStream; out Value: UInt64; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUInt64_BE(Stream,Value)
else
  Result := Stream_ReadUInt64_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt64_LE(Stream: TStream; Advance: Boolean = True): UInt64;
begin
Stream_ReadUInt64_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt64_BE(Stream: TStream; Advance: Boolean = True): UInt64;
begin
Stream_ReadUInt64_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUInt64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUInt64_BE(Stream,Advance)
else
  Result := Stream_GetUInt64_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUInt64(Stream: TStream; Endian: TEndian = endDefault): UInt64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUInt64_BE(Stream)
else
  Result := Stream_GetUInt64_LE(Stream);
end;

{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Stream_ReadFloat32_LE(Stream: TStream; out Value: Float32; Advance: Boolean = True): TMemSize;
var
  Temp: UInt32 absolute Value;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadFloat32_BE(Stream: TStream; out Value: Float32; Advance: Boolean = True): TMemSize;
var
  Temp: UInt32 absolute Value;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadFloat32(Stream: TStream; out Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadFloat32_BE(Stream,Value,Advance)
else
  Result := Stream_ReadFloat32_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadFloat32(Stream: TStream; out Value: Float32; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadFloat32_BE(Stream,Value)
else
  Result := Stream_ReadFloat32_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat32_LE(Stream: TStream; Advance: Boolean = True): Float32;
begin
Stream_ReadFloat32_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat32_BE(Stream: TStream; Advance: Boolean = True): Float32;
begin
Stream_ReadFloat32_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetFloat32_BE(Stream,Advance)
else
  Result := Stream_GetFloat32_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetFloat32(Stream: TStream; Endian: TEndian = endDefault): Float32;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetFloat32_BE(Stream)
else
  Result := Stream_GetFloat32_LE(Stream);
end;

//==============================================================================

Function Stream_ReadFloat64_LE(Stream: TStream; out Value: Float64; Advance: Boolean = True): TMemSize;
var
  Temp: UInt64 absolute Value;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadFloat64_BE(Stream: TStream; out Value: Float64; Advance: Boolean = True): TMemSize;
var
  Temp: UInt64 absolute Value;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadFloat64(Stream: TStream; out Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadFloat64_BE(Stream,Value,Advance)
else
  Result := Stream_ReadFloat64_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadFloat64(Stream: TStream; out Value: Float64; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadFloat64_BE(Stream,Value)
else
  Result := Stream_ReadFloat64_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat64_LE(Stream: TStream; Advance: Boolean = True): Float64;
begin
Stream_ReadFloat64_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat64_BE(Stream: TStream; Advance: Boolean = True): Float64;
begin
Stream_ReadFloat64_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetFloat64_BE(Stream,Advance)
else
  Result := Stream_GetFloat64_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetFloat64(Stream: TStream; Endian: TEndian = endDefault): Float64;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetFloat64_BE(Stream)
else
  Result := Stream_GetFloat64_LE(Stream);
end;

//==============================================================================

Function Stream_ReadFloat80_LE(Stream: TStream; out Value: Float80; Advance: Boolean = True): TMemSize;
var
  Temp: TFloat80Overlay absolute Value;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadFloat80_BE(Stream: TStream; out Value: Float80; Advance: Boolean = True): TMemSize;
var
  Temp: TFloat80Overlay absolute Value;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadFloat80(Stream: TStream; out Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadFloat80_BE(Stream,Value,Advance)
else
  Result := Stream_ReadFloat80_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadFloat80(Stream: TStream; out Value: Float80; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadFloat80_BE(Stream,Value)
else
  Result := Stream_ReadFloat80_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat80_LE(Stream: TStream; Advance: Boolean = True): Float80;
begin
Stream_ReadFloat80_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat80_BE(Stream: TStream; Advance: Boolean = True): Float80;
begin
Stream_ReadFloat80_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetFloat80(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float80;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetFloat80_BE(Stream,Advance)
else
  Result := Stream_GetFloat80_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetFloat80(Stream: TStream; Endian: TEndian = endDefault): Float80;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetFloat80_BE(Stream)
else
  Result := Stream_GetFloat80_LE(Stream);
end;

//==============================================================================

Function Stream_ReadDateTime_LE(Stream: TStream; out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := Stream_ReadFloat64_LE(Stream,Float64(Value),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadDateTime_BE(Stream: TStream; out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := Stream_ReadFloat64_BE(Stream,Float64(Value),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadDateTime(Stream: TStream; out Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_ReadFloat64(Stream,Float64(Value),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadDateTime(Stream: TStream; out Value: TDateTime; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_ReadFloat64(Stream,Float64(Value),Endian);
end;

//------------------------------------------------------------------------------

Function Stream_GetDateTime_LE(Stream: TStream; Advance: Boolean = True): TDateTime;
begin
Result := Stream_GetFloat64_LE(Stream,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetDateTime_BE(Stream: TStream; Advance: Boolean = True): TDateTime;
begin
Result := Stream_GetFloat64_BE(Stream,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetDateTime(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): TDateTime;
begin
Result := Stream_GetFloat64(Stream,Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetDateTime(Stream: TStream; Endian: TEndian = endDefault): TDateTime;
begin
Result := Stream_GetFloat64(Stream,Endian);
end;

//==============================================================================

Function Stream_ReadCurrency_LE(Stream: TStream; out Value: Currency; Advance: Boolean = True): TMemSize;
var
  Temp: UInt64 absolute Value;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
{$IFDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadCurrency_BE(Stream: TStream; out Value: Currency; Advance: Boolean = True): TMemSize;
var
  Temp: UInt64 absolute Value;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
{$IFNDEF ENDIAN_BIG}
Temp := SwapEndian(Temp);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadCurrency(Stream: TStream; out Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadCurrency_BE(Stream,Value,Advance)
else
  Result := Stream_ReadCurrency_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadCurrency(Stream: TStream; out Value: Currency; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadCurrency_BE(Stream,Value)
else
  Result := Stream_ReadCurrency_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetCurrency_LE(Stream: TStream; Advance: Boolean = True): Currency;
begin
Stream_ReadCurrency_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetCurrency_BE(Stream: TStream; Advance: Boolean = True): Currency;
begin
Stream_ReadCurrency_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetCurrency(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Currency;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetCurrency_BE(Stream,Advance)
else
  Result := Stream_GetCurrency_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetCurrency(Stream: TStream; Endian: TEndian = endDefault): Currency;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetCurrency_BE(Stream)
else
  Result := Stream_GetCurrency_LE(Stream);
end;

{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function _Stream_ReadAnsiChar(Stream: TStream; out Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadAnsiChar_LE(Stream: TStream; out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadAnsiChar(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadAnsiChar_BE(Stream: TStream; out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadAnsiChar(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadAnsiChar(Stream: TStream; out Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadAnsiChar_BE(Stream,Value,Advance)
else
  Result := Stream_ReadAnsiChar_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadAnsiChar(Stream: TStream; out Value: AnsiChar; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadAnsiChar_BE(Stream,Value)
else
  Result := Stream_ReadAnsiChar_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetAnsiChar_LE(Stream: TStream; Advance: Boolean = True): AnsiChar;
begin
_Stream_ReadAnsiChar(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetAnsiChar_BE(Stream: TStream; Advance: Boolean = True): AnsiChar;
begin
_Stream_ReadAnsiChar(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetAnsiChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): AnsiChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetAnsiChar_BE(Stream,Advance)
else
  Result := Stream_GetAnsiChar_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetAnsiChar(Stream: TStream; Endian: TEndian = endDefault): AnsiChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetAnsiChar_BE(Stream)
else
  Result := Stream_GetAnsiChar_LE(Stream);
end;

//==============================================================================

Function _Stream_ReadUTF8Char(Stream: TStream; out Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUTF8Char_LE(Stream: TStream; out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadUTF8Char(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUTF8Char_BE(Stream: TStream; out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadUTF8Char(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUTF8Char(Stream: TStream; out Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUTF8Char_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUTF8Char_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUTF8Char(Stream: TStream; out Value: UTF8Char; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUTF8Char_BE(Stream,Value)
else
  Result := Stream_ReadUTF8Char_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetUTF8Char_LE(Stream: TStream; Advance: Boolean = True): UTF8Char;
begin
_Stream_ReadUTF8Char(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUTF8Char_BE(Stream: TStream; Advance: Boolean = True): UTF8Char;
begin
_Stream_ReadUTF8Char(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUTF8Char(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UTF8Char;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUTF8Char_BE(Stream,Advance)
else
  Result := Stream_GetUTF8Char_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUTF8Char(Stream: TStream; Endian: TEndian = endDefault): UTF8Char;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUTF8Char_BE(Stream)
else
  Result := Stream_GetUTF8Char_LE(Stream);
end;
 
//==============================================================================

Function Stream_ReadWideChar_LE(Stream: TStream; out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := WideChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadWideChar_BE(Stream: TStream; out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := WideChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadWideChar(Stream: TStream; out Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadWideChar_BE(Stream,Value,Advance)
else
  Result := Stream_ReadWideChar_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadWideChar(Stream: TStream; out Value: WideChar; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadWideChar_BE(Stream,Value)
else
  Result := Stream_ReadWideChar_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetWideChar_LE(Stream: TStream; Advance: Boolean = True): WideChar;
begin
Stream_ReadWideChar_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetWideChar_BE(Stream: TStream; Advance: Boolean = True): WideChar;
begin
Stream_ReadWideChar_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetWideChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): WideChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetWideChar_BE(Stream,Advance)
else
  Result := Stream_GetWideChar_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetWideChar(Stream: TStream; Endian: TEndian = endDefault): WideChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetWideChar_BE(Stream)
else
  Result := Stream_GetWideChar_LE(Stream);
end;

//==============================================================================

Function Stream_ReadUnicodeChar_LE(Stream: TStream; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := UnicodeChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUnicodeChar_BE(Stream: TStream; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := UnicodeChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUnicodeChar(Stream: TStream; out Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUnicodeChar_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUnicodeChar_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUnicodeChar(Stream: TStream; out Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUnicodeChar_BE(Stream,Value)
else
  Result := Stream_ReadUnicodeChar_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetUnicodeChar_LE(Stream: TStream; Advance: Boolean = True): UnicodeChar;
begin
Stream_ReadUnicodeChar_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUnicodeChar_BE(Stream: TStream; Advance: Boolean = True): UnicodeChar;
begin
Stream_ReadUnicodeChar_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUnicodeChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UnicodeChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUnicodeChar_BE(Stream,Advance)
else
  Result := Stream_GetUnicodeChar_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUnicodeChar(Stream: TStream; Endian: TEndian = endDefault): UnicodeChar;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUnicodeChar_BE(Stream)
else
  Result := Stream_GetUnicodeChar_LE(Stream);
end;

//==============================================================================

Function Stream_ReadUCS4Char_LE(Stream: TStream; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := UCS4Char(SwapEndian(UInt32(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUCS4Char_BE(Stream: TStream; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFNDEF ENDIAN_BIG}
Value := UCS4Char(SwapEndian(UInt32(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUCS4Char(Stream: TStream; out Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUCS4Char_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUCS4Char_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUCS4Char(Stream: TStream; out Value: UCS4Char; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUCS4Char_BE(Stream,Value)
else
  Result := Stream_ReadUCS4Char_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetUCS4Char_LE(Stream: TStream; Advance: Boolean = True): UCS4Char;
begin
Stream_ReadUCS4Char_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUCS4Char_BE(Stream: TStream; Advance: Boolean = True): UCS4Char;
begin
Stream_ReadUCS4Char_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUCS4Char(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UCS4Char;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUCS4Char_BE(Stream,Advance)
else
  Result := Stream_GetUCS4Char_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUCS4Char(Stream: TStream; Endian: TEndian = endDefault): UCS4Char;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUCS4Char_BE(Stream)
else
  Result := Stream_GetUCS4Char_LE(Stream);
end;
 
//==============================================================================

Function Stream_ReadChar_LE(Stream: TStream; out Value: Char; Advance: Boolean = True): TMemSize;
var
  Temp: UInt16;
begin
Result := Stream_ReadUInt16_LE(Stream,Temp,Advance);
Value := Char(Temp);
end;

//------------------------------------------------------------------------------

Function Stream_ReadChar_BE(Stream: TStream; out Value: Char; Advance: Boolean = True): TMemSize;
var
  Temp: UInt16;
begin
Result := Stream_ReadUInt16_BE(Stream,Temp,Advance);
Value := Char(Temp);
end;

//------------------------------------------------------------------------------

Function Stream_ReadChar(Stream: TStream; out Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  Temp: UInt16;
begin
Result := Stream_ReadUInt16(Stream,Temp,Advance,Endian);
Value := Char(Temp);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadChar(Stream: TStream; out Value: Char; Endian: TEndian = endDefault): TMemSize;
var
  Temp: UInt16;
begin
Result := Stream_ReadUInt16(Stream,Temp,Endian);
Value := Char(Temp);
end;

//------------------------------------------------------------------------------

Function Stream_GetChar_LE(Stream: TStream; Advance: Boolean = True): Char;
begin
Result := Char(Stream_GetUInt16_LE(Stream,Advance));
end;

//------------------------------------------------------------------------------

Function Stream_GetChar_BE(Stream: TStream; Advance: Boolean = True): Char;
begin
Result := Char(Stream_GetUInt16_BE(Stream,Advance));
end;

//------------------------------------------------------------------------------

Function Stream_GetChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Char;
begin
Result := Char(Stream_GetUInt16(Stream,Advance,Endian));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetChar(Stream: TStream; Endian: TEndian = endDefault): Char;
begin
Result := Char(Stream_GetUInt16(Stream,Endian));
end;

{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function _Stream_ReadShortString(Stream: TStream; out Str: ShortString; Advance: Boolean): TMemSize;
var
  StrLength:  UInt8;
begin
Result := Stream_ReadUInt8_LE(Stream,StrLength,True);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_LE(Stream,Addr(Str[1])^,StrLength,True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadShortString_LE(Stream: TStream; out Str: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadShortString(Stream,Str,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadShortString_BE(Stream: TStream; out Str: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadShortString(Stream,Str,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadShortString(Stream: TStream; out Str: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadShortString_BE(Stream,Str,Advance)
else
  Result := Stream_ReadShortString_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadShortString(Stream: TStream; out Str: ShortString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadShortString_BE(Stream,Str)
else
  Result := Stream_ReadShortString_LE(Stream,Str);
end;

//------------------------------------------------------------------------------

Function Stream_GetShortString_LE(Stream: TStream; Advance: Boolean = True): ShortString;
begin
_Stream_ReadShortString(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetShortString_BE(Stream: TStream; Advance: Boolean = True): ShortString;
begin
_Stream_ReadShortString(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetShortString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): ShortString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetShortString_BE(Stream,Advance)
else
  Result := Stream_GetShortString_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetShortString(Stream: TStream; Endian: TEndian = endDefault): ShortString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetShortString_BE(Stream)
else
  Result := Stream_GetShortString_LE(Stream);
end;

//==============================================================================

Function Stream_ReadAnsiString_LE(Stream: TStream; out Str: AnsiString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_LE(Stream,PAnsiChar(Str)^,StrLength * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadAnsiString_BE(Stream: TStream; out Str: AnsiString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_BE(Stream,PAnsiChar(Str)^,StrLength * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadAnsiString(Stream: TStream; out Str: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadAnsiString_BE(Stream,Str,Advance)
else
  Result := Stream_ReadAnsiString_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadAnsiString(Stream: TStream; out Str: AnsiString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadAnsiString_BE(Stream,Str)
else
  Result := Stream_ReadAnsiString_LE(Stream,Str);
end;

//------------------------------------------------------------------------------

Function Stream_GetAnsiString_LE(Stream: TStream; Advance: Boolean = True): AnsiString;
begin
Stream_ReadAnsiString_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetAnsiString_BE(Stream: TStream; Advance: Boolean = True): AnsiString;
begin
Stream_ReadAnsiString_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetAnsiString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): AnsiString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetAnsiString_BE(Stream,Advance)
else
  Result := Stream_GetAnsiString_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetAnsiString(Stream: TStream; Endian: TEndian = endDefault): AnsiString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetAnsiString_BE(Stream)
else
  Result := Stream_GetAnsiString_LE(Stream);
end;

//==============================================================================

Function Stream_ReadUTF8String_LE(Stream: TStream; out Str: UTF8String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_LE(Stream,PUTF8Char(Str)^,StrLength * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUTF8String_BE(Stream: TStream; out Str: UTF8String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_BE(Stream,PUTF8Char(Str)^,StrLength * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUTF8String(Stream: TStream; out Str: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUTF8String_BE(Stream,Str,Advance)
else
  Result := Stream_ReadUTF8String_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUTF8String(Stream: TStream; out Str: UTF8String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUTF8String_BE(Stream,Str)
else
  Result := Stream_ReadUTF8String_LE(Stream,Str);
end;

//------------------------------------------------------------------------------

Function Stream_GetUTF8String_LE(Stream: TStream; Advance: Boolean = True): UTF8String;
begin
Stream_ReadUTF8String_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUTF8String_BE(Stream: TStream; Advance: Boolean = True): UTF8String;
begin
Stream_ReadUTF8String_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUTF8String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UTF8String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUTF8String_BE(Stream,Advance)
else
  Result := Stream_GetUTF8String_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUTF8String(Stream: TStream; Endian: TEndian = endDefault): UTF8String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUTF8String_BE(Stream)
else
  Result := Stream_GetUTF8String_LE(Stream);
end;

//==============================================================================

Function Stream_ReadWideString_LE(Stream: TStream; out Str: WideString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Str)),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,PWideChar(Str)^,StrLength * SizeOf(WideChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadWideString_BE(Stream: TStream; out Str: WideString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,PWideChar(Str)^,StrLength * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Str)),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadWideString(Stream: TStream; out Str: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadWideString_BE(Stream,Str,Advance)
else
  Result := Stream_ReadWideString_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadWideString(Stream: TStream; out Str: WideString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadWideString_BE(Stream,Str)
else
  Result := Stream_ReadWideString_LE(Stream,Str);
end;

//------------------------------------------------------------------------------

Function Stream_GetWideString_LE(Stream: TStream; Advance: Boolean = True): WideString;
begin
Stream_ReadWideString_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetWideString_BE(Stream: TStream; Advance: Boolean = True): WideString;
begin
Stream_ReadWideString_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetWideString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): WideString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetWideString_BE(Stream,Advance)
else
  Result := Stream_GetWideString_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetWideString(Stream: TStream; Endian: TEndian = endDefault): WideString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetWideString_BE(Stream)
else
  Result := Stream_GetWideString_LE(Stream);
end;

//==============================================================================

Function Stream_ReadUnicodeString_LE(Stream: TStream; out Str: UnicodeString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Str)),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,PUnicodeChar(Str)^,StrLength * SizeOf(UnicodeChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUnicodeString_BE(Stream: TStream; out Str: UnicodeString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,PUnicodeChar(Str)^,StrLength * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Str)),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUnicodeString(Stream: TStream; out Str: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUnicodeString_BE(Stream,Str,Advance)
else
  Result := Stream_ReadUnicodeString_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUnicodeString(Stream: TStream; out Str: UnicodeString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUnicodeString_BE(Stream,Str)
else
  Result := Stream_ReadUnicodeString_LE(Stream,Str);
end;

//------------------------------------------------------------------------------

Function Stream_GetUnicodeString_LE(Stream: TStream; Advance: Boolean = True): UnicodeString;
begin
Stream_ReadUnicodeString_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUnicodeString_BE(Stream: TStream; Advance: Boolean = True): UnicodeString;
begin
Stream_ReadUnicodeString_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUnicodeString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUnicodeString_BE(Stream,Advance)
else
  Result := Stream_GetUnicodeString_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUnicodeString(Stream: TStream; Endian: TEndian = endDefault): UnicodeString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUnicodeString_BE(Stream)
else
  Result := Stream_GetUnicodeString_LE(Stream);
end;

//==============================================================================

Function Stream_ReadUCS4String_LE(Stream: TStream; out Str: UCS4String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength + 1);
Str[High(Str)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Str[Low(Str)])),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,Addr(Str[Low(Str)])^,StrLength * SizeOf(UCS4Char),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUCS4String_BE(Stream: TStream; out Str: UCS4String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
SetLength(Str,StrLength + 1);
Str[High(Str)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,Addr(Str[Low(Str)])^,StrLength * SizeOf(UCS4Char),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Str[Low(Str)])),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUCS4String(Stream: TStream; out Str: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUCS4String_BE(Stream,Str,Advance)
else
  Result := Stream_ReadUCS4String_LE(Stream,Str,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUCS4String(Stream: TStream; out Str: UCS4String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUCS4String_BE(Stream,Str)
else
  Result := Stream_ReadUCS4String_LE(Stream,Str);
end;

//------------------------------------------------------------------------------

Function Stream_GetUCS4String_LE(Stream: TStream; Advance: Boolean = True): UCS4String;
begin
Stream_ReadUCS4String_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUCS4String_BE(Stream: TStream; Advance: Boolean = True): UCS4String;
begin
Stream_ReadUCS4String_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetUCS4String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UCS4String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUCS4String_BE(Stream,Advance)
else
  Result := Stream_GetUCS4String_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetUCS4String(Stream: TStream; Endian: TEndian = endDefault): UCS4String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetUCS4String_BE(Stream)
else
  Result := Stream_GetUCS4String_LE(Stream);
end;

//==============================================================================

Function Stream_ReadString_LE(Stream: TStream; out Str: String; Advance: Boolean = True): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadUTF8String_LE(Stream,TempStr,Advance);
Str := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Stream_ReadString_BE(Stream: TStream; out Str: String; Advance: Boolean = True): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadUTF8String_BE(Stream,TempStr,Advance);
Str := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Stream_ReadString(Stream: TStream; out Str: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadUTF8String(Stream,TempStr,Advance,Endian);
Str := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadString(Stream: TStream; out Str: String; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadUTF8String(Stream,TempStr,Endian);
Str := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Stream_GetString_LE(Stream: TStream; Advance: Boolean = True): String;
begin
Result := UTF8ToStr(Stream_GetUTF8String_LE(Stream,Advance));
end;

//------------------------------------------------------------------------------

Function Stream_GetString_BE(Stream: TStream; Advance: Boolean = True): String;
begin
Result := UTF8ToStr(Stream_GetUTF8String_BE(Stream,Advance));
end;

//------------------------------------------------------------------------------

Function Stream_GetString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): String;
begin
Result := UTF8ToStr(Stream_GetUTF8String(Stream,Advance,Endian));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetString(Stream: TStream; Endian: TEndian = endDefault): String;
begin
Result := UTF8ToStr(Stream_GetUTF8String(Stream,Endian));
end;

{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function _Stream_ReadBuffer(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Buffer)^,Size);
Result := Size;
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadBuffer_LE(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadBuffer(Stream,Buffer,Size,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadBuffer_BE(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadBuffer(Stream,Buffer,Size,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadBuffer(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadBuffer_BE(Stream,Buffer,Size,Advance)
else
  Result := Stream_ReadBuffer_LE(Stream,Buffer,Size,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadBuffer(Stream: TStream; out Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadBuffer_BE(Stream,Buffer,Size)
else
  Result := Stream_ReadBuffer_LE(Stream,Buffer,Size);
end;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Stream_ReadVariant_LE(Stream: TStream; out Value: Variant; Advance: Boolean = True): TMemSize;
var
  VariantTypeInt: UInt8;
  Dimensions:     Integer;
  i:              Integer;
  IndicesBounds:  array of PtrInt;
  Indices:        array of Integer;
  WideStrTemp:    WideString;
  AnsiStrTemp:    AnsiString;
  UnicodeStrTemp: UnicodeString;

  Function ReadVarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:    Integer;
    TempVar:  Variant;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          begin
            Inc(Result,Stream_ReadVariant_LE(Stream,TempVar,True));
            VarArrayPut(Value,TempVar,Indices);
          end
        else Inc(Result,ReadVarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
VariantTypeInt := Stream_GetUInt8_LE(Stream,True);
If VariantTypeInt and BS_VARTYPENUM_ARRAY <> 0 then
  begin
    Dimensions := Stream_GetInt32_LE(Stream,True);
    Result := 5 + (Dimensions * 8);
    If Dimensions > 0 then
      begin
        SetLength(IndicesBounds,Dimensions * 2);
        For i := Low(IndicesBounds) to High(IndicesBounds) do
          IndicesBounds[i] := PtrInt(Stream_GetInt32_LE(Stream,True));
        Value := VarArrayCreate(IndicesBounds,IntToVarType(VariantTypeInt and BS_VARTYPENUM_TYPEMASK));
        SetLength(Indices,Dimensions);
        Inc(Result,ReadVarArrayDimension(1));
      end;
  end
else
  begin
    // simple type
    TVarData(Value).vType := IntToVarType(VariantTypeInt);
    case VariantTypeInt and BS_VARTYPENUM_TYPEMASK of
      BS_VARTYPENUM_BOOLEN:   begin
                                TVarData(Value).vBoolean := Stream_GetBool_LE(Stream,True);
                                Result := StreamedSize_Bool + 1;
                              end;
      BS_VARTYPENUM_SHORTINT: Result := Stream_ReadInt8_LE(Stream,TVarData(Value).vShortInt,True) + 1;
      BS_VARTYPENUM_SMALLINT: Result := Stream_ReadInt16_LE(Stream,TVarData(Value).vSmallInt,True) + 1;
      BS_VARTYPENUM_INTEGER:  Result := Stream_ReadInt32_LE(Stream,TVarData(Value).vInteger,True) + 1;
      BS_VARTYPENUM_INT64:    Result := Stream_ReadInt64_LE(Stream,TVarData(Value).vInt64,True) + 1;
      BS_VARTYPENUM_BYTE:     Result := Stream_ReadUInt8_LE(Stream,TVarData(Value).vByte,True) + 1;
      BS_VARTYPENUM_WORD:     Result := Stream_ReadUInt16_LE(Stream,TVarData(Value).vWord,True) + 1;
      BS_VARTYPENUM_LONGWORD: Result := Stream_ReadUInt32_LE(Stream,TVarData(Value).vLongWord,True) + 1;
      BS_VARTYPENUM_UINT64: {$IF Declared(varUInt64)}
                              Result := Stream_ReadUInt64_LE(Stream,TVarData(Value).{$IFDEF FPC}vQWord{$ELSE}vUInt64{$ENDIF},True) + 1;
                            {$ELSE}
                              Result := Stream_ReadUInt64_LE(Stream,UInt64(TVarData(Value).vInt64),True) + 1;
                            {$IFEND}
      BS_VARTYPENUM_SINGLE:   Result := Stream_ReadFloat32_LE(Stream,TVarData(Value).vSingle,True) + 1;
      BS_VARTYPENUM_DOUBLE:   Result := Stream_ReadFloat64_LE(Stream,TVarData(Value).vDouble,True) + 1;
      BS_VARTYPENUM_CURRENCY: Result := Stream_ReadCurrency_LE(Stream,TVarData(Value).vCurrency,True) + 1;
      BS_VARTYPENUM_DATE:     Result := Stream_ReadDateTime_LE(Stream,TVarData(Value).vDate,True) + 1;
      BS_VARTYPENUM_OLESTR:   begin
                                Result := Stream_ReadWideString_LE(Stream,WideStrTemp,True) + 1;
                                Value := VarAsType(WideStrTemp,varOleStr);
                              end;
      BS_VARTYPENUM_STRING:   begin
                                Result := Stream_ReadAnsiString_LE(Stream,AnsiStrTemp,True) + 1;
                                Value := VarAsType(AnsiStrTemp,varString);
                              end;
      BS_VARTYPENUM_USTRING:  begin
                                Result := Stream_ReadUnicodeString_LE(Stream,UnicodeStrTemp,True) + 1;
                              {$IF Declared(varUString)}
                                Value := VarAsType(UnicodeStrTemp,varUString);
                              {$ELSE}
                                Value := VarAsType(UnicodeStrTemp,varOleStr);
                              {$IFEND}
                              end;
    else
      raise EBSUnsupportedVarType.CreateFmt('Stream_ReadVariant_LE: Cannot read variant of this type number (%d).',[VariantTypeInt and BS_VARTYPENUM_TYPEMASK]);
    end;
  end;
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadVariant_BE(Stream: TStream; out Value: Variant; Advance: Boolean = True): TMemSize;
var
  VariantTypeInt: UInt8;
  Dimensions:     Integer;
  i:              Integer;
  IndicesBounds:  array of PtrInt;
  Indices:        array of Integer;
  WideStrTemp:    WideString;
  AnsiStrTemp:    AnsiString;
  UnicodeStrTemp: UnicodeString;

  Function ReadVarArrayDimension(Dimension: Integer): TMemSize;
  var
    Index:    Integer;
    TempVar:  Variant;
  begin
    Result := 0;
    For Index := VarArrayLowBound(Value,Dimension) to VarArrayHighBound(Value,Dimension) do
      begin
        Indices[Pred(Dimension)] := Index;
        If Dimension >= Dimensions then
          begin
            Inc(Result,Stream_ReadVariant_BE(Stream,TempVar,True));
            VarArrayPut(Value,TempVar,Indices);
          end
        else Inc(Result,ReadVarArrayDimension(Succ(Dimension)));
      end;
  end;

begin
VariantTypeInt := Stream_GetUInt8_BE(Stream,True);
If VariantTypeInt and BS_VARTYPENUM_ARRAY <> 0 then
  begin
    Dimensions := Stream_GetInt32_BE(Stream,True);
    Result := 5 + (Dimensions * 8);
    If Dimensions > 0 then
      begin
        SetLength(IndicesBounds,Dimensions * 2);
        For i := Low(IndicesBounds) to High(IndicesBounds) do
          IndicesBounds[i] := PtrInt(Stream_GetInt32_BE(Stream,True));
        Value := VarArrayCreate(IndicesBounds,IntToVarType(VariantTypeInt and BS_VARTYPENUM_TYPEMASK));
        SetLength(Indices,Dimensions);
        Inc(Result,ReadVarArrayDimension(1));
      end;
  end
else
  begin
    // simple type
    TVarData(Value).vType := IntToVarType(VariantTypeInt);
    case VariantTypeInt and BS_VARTYPENUM_TYPEMASK of
      BS_VARTYPENUM_BOOLEN:   begin
                                TVarData(Value).vBoolean := Stream_GetBool_BE(Stream,True);
                                Result := StreamedSize_Bool + 1;
                              end;
      BS_VARTYPENUM_SHORTINT: Result := Stream_ReadInt8_BE(Stream,TVarData(Value).vShortInt,True) + 1;
      BS_VARTYPENUM_SMALLINT: Result := Stream_ReadInt16_BE(Stream,TVarData(Value).vSmallInt,True) + 1;
      BS_VARTYPENUM_INTEGER:  Result := Stream_ReadInt32_BE(Stream,TVarData(Value).vInteger,True) + 1;
      BS_VARTYPENUM_INT64:    Result := Stream_ReadInt64_BE(Stream,TVarData(Value).vInt64,True) + 1;
      BS_VARTYPENUM_BYTE:     Result := Stream_ReadUInt8_BE(Stream,TVarData(Value).vByte,True) + 1;
      BS_VARTYPENUM_WORD:     Result := Stream_ReadUInt16_BE(Stream,TVarData(Value).vWord,True) + 1;
      BS_VARTYPENUM_LONGWORD: Result := Stream_ReadUInt32_BE(Stream,TVarData(Value).vLongWord,True) + 1;
      BS_VARTYPENUM_UINT64: {$IF Declared(varUInt64)}
                              Result := Stream_ReadUInt64_BE(Stream,TVarData(Value).{$IFDEF FPC}vQWord{$ELSE}vUInt64{$ENDIF},True) + 1;
                            {$ELSE}
                              Result := Stream_ReadUInt64_BE(Stream,UInt64(TVarData(Value).vInt64),True) + 1;
                            {$IFEND}
      BS_VARTYPENUM_SINGLE:   Result := Stream_ReadFloat32_BE(Stream,TVarData(Value).vSingle,True) + 1;
      BS_VARTYPENUM_DOUBLE:   Result := Stream_ReadFloat64_BE(Stream,TVarData(Value).vDouble,True) + 1;
      BS_VARTYPENUM_CURRENCY: Result := Stream_ReadCurrency_BE(Stream,TVarData(Value).vCurrency,True) + 1;
      BS_VARTYPENUM_DATE:     Result := Stream_ReadDateTime_BE(Stream,TVarData(Value).vDate,True) + 1;
      BS_VARTYPENUM_OLESTR:   begin
                                Result := Stream_ReadWideString_BE(Stream,WideStrTemp,True) + 1;
                                Value := VarAsType(WideStrTemp,varOleStr);
                              end;
      BS_VARTYPENUM_STRING:   begin
                                Result := Stream_ReadAnsiString_BE(Stream,AnsiStrTemp,True) + 1;
                                Value := VarAsType(AnsiStrTemp,varString);
                              end;
      BS_VARTYPENUM_USTRING:  begin
                                Result := Stream_ReadUnicodeString_BE(Stream,UnicodeStrTemp,True) + 1;
                              {$IF Declared(varUString)}
                                Value := VarAsType(UnicodeStrTemp,varUString);
                              {$ELSE}
                                Value := VarAsType(UnicodeStrTemp,varOleStr);
                              {$IFEND}
                              end;
    else
      raise EBSUnsupportedVarType.CreateFmt('Stream_ReadVariant_BE: Cannot read variant of this type number (%d).',[VariantTypeInt and BS_VARTYPENUM_TYPEMASK]);
    end;
  end;
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadVariant(Stream: TStream; out Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadVariant_BE(Stream,Value,Advance)
else
  Result := Stream_ReadVariant_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadVariant(Stream: TStream; out Value: Variant; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadVariant_BE(Stream,Value)
else
  Result := Stream_ReadVariant_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetVariant_LE(Stream: TStream; Advance: Boolean = True): Variant;
begin
Stream_ReadVariant_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetVariant_BE(Stream: TStream; Advance: Boolean = True): Variant;
begin
Stream_ReadVariant_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetVariant(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Variant;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetVariant_BE(Stream,Advance)
else
  Result := Stream_GetVariant_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetVariant(Stream: TStream; Endian: TEndian = endDefault): Variant;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetVariant_BE(Stream)
else
  Result := Stream_GetVariant_LE(Stream);
end;


{===============================================================================
--------------------------------------------------------------------------------
                                 TCustomStreamer
--------------------------------------------------------------------------------
===============================================================================}
{===============================================================================
    TCustomStreamer - class implementation
===============================================================================}
{-------------------------------------------------------------------------------
    TCustomStreamer - protected methods
-------------------------------------------------------------------------------}
(*
Function TCustomStreamer.GetCapacity: Integer;
begin
Result := Length(fBookmarks);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.SetCapacity(Value: Integer);
begin
If Value <> Length(fBookmarks) then
  begin
    SetLength(fBookMarks,Value);
    If Value < fCount then
      fCount := Value;
  end;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCount: Integer;
begin
Result := fCount;
end;

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W5024{$ENDIF}
procedure TCustomStreamer.SetCount(Value: Integer);
begin
// do nothing
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

Function TCustomStreamer.GetBookmark(Index: Integer): Int64;
begin
If CheckIndex(Index) then
  Result := fBookmarks[Index]
else
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.GetBookmark: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.SetBookmark(Index: Integer; Value: Int64);
begin
If CheckIndex(Index) then
  fBookmarks[Index] := Value
else
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.SetBookmark: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetDistance: Int64;
begin
Result := CurrentPosition - StartPosition;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.Initialize;
begin
SetLength(fBookmarks,0);
fCount := 0;
fStartPosition := 0;
fEndian := endDefault;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.Finalize;
begin
// nothing to do
end;

{-------------------------------------------------------------------------------
    TCustomStreamer - public methods
-------------------------------------------------------------------------------}

destructor TCustomStreamer.Destroy;
begin
Finalize;
inherited;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LowIndex: Integer;
begin
Result := Low(fBookmarks);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.HighIndex: Integer;
begin
Result := Pred(fCount);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.MoveToStart;
begin
CurrentPosition := StartPosition;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.MoveToBookmark(Index: Integer);
begin
If CheckIndex(Index) then
  CurrentPosition := fBookmarks[Index]
else
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.MoveToBookmark: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.MoveBy(Offset: Int64);
begin
CurrentPosition := CurrentPosition + Offset;
end;

//------------------------------------------------------------------------------
 
Function TCustomStreamer.IndexOfBookmark(Position: Int64): Integer;
var
  i:  Integer;
begin
Result := -1;
For i := LowIndex to HighIndex do
  If fBookmarks[i] = Position then
    begin
      Result := i;
      Break;
    end;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.AddBookmark: Integer;
begin
Result := AddBookmark(CurrentPosition);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.AddBookmark(Position: Int64): Integer;
begin
Grow;
Result := fCount;
fBookmarks[Result] := Position;
Inc(fCount);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.RemoveBookmark(Position: Int64; RemoveAll: Boolean = True): Integer;
begin
repeat
  Result := IndexOfBookmark(Position);
  If CheckIndex(Result) then
    DeleteBookMark(Result);
until not CheckIndex(Result) or not RemoveAll;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.DeleteBookmark(Index: Integer);
var
  i:  Integer;
begin
If CheckIndex(Index) then
  begin
    For i := Index to Pred(High(fBookmarks)) do
      fBookmarks[i] := fBookMarks[i + 1];
    Dec(fCount);
    Shrink;
  end
else raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.DeleteBookmark: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.ClearBookmark;
begin
fCount := 0;
Shrink;
end;

//==============================================================================

Function TCustomStreamer.WriteBool(Value: ByteBool; Advance: Boolean = True): TMemSize;
var
  Temp: UInt8;
begin
Temp := BoolToNum(Value);
Result := WriteValue(@Temp,Advance,SizeOf(Temp),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBoolean(Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := WriteBool(Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt8(Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt8(Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive1B);
end;
 
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt16(Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt16(Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt32(Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt32(Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt64(Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt64(Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat32(Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat64(Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat80(Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive10B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteDateTime(Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCurrency(Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;   

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiChar(Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8Char(Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideChar(Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeChar(Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4Char(Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,SizeOf(Value),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteChar(Value: Char; Advance: Boolean = True): TMemSize;
var
  Temp: UInt16;
begin
Temp := UInt16(Ord(Value));
Result := WriteValue(@Temp,Advance,SizeOf(Temp),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteShortString(const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,0,vtShortString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiString(const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,0,vtAnsiString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8String(const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,0,vtUTF8String);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideString(const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,0,vtWideString);
end;
 
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeString(const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,0,vtUnicodeString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4String(const Value: UCS4String; Advance: Boolean = True): TMemSize; 
begin
Result := WriteValue(@Value,Advance,0,vtUCS4String);
end;
  
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteString(const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,0,vtString);
end;
  
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBuffer(const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Buffer,Advance,Size,vtBytes);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBytes(const Value: array of UInt8; Advance: Boolean = True): TMemSize;
var
  i:      Integer;
  OldPos: Int64;
begin
OldPos := CurrentPosition;
Result := 0;
For i := Low(Value) to High(Value) do
  WriteUInt8(Value[i],True);
If not Advance then
  CurrentPosition := OldPos;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.FillBytes(ByteCount: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,ByteCount,vtFillBytes);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteVariant(const Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(@Value,Advance,0,vtVariant);
end;

//==============================================================================

Function TCustomStreamer.ReadBool(out Value: ByteBool; Advance: Boolean = True): TMemSize;
var
  Temp: UInt8;
begin
Result := ReadValue(@Temp,Advance,SizeOf(Temp),vtPrimitive1B);
Value := NumToBool(Temp);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadBool(Advance: Boolean = True): ByteBool;
var
  Temp: UInt8;
begin
ReadValue(@Temp,Advance,SizeOf(Temp),vtPrimitive1B);
Result := NumToBool(Temp);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBoolean(out Value: Boolean; Advance: Boolean = True): TMemSize;
var
  TempBool: ByteBool;
begin
Result := ReadBool(TempBool,Advance);
Value := TempBool;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadBoolean(Advance: Boolean = True): Boolean;
begin
Result := LoadBool(Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt8(out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadInt8(Advance: Boolean = True): Int8;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt8(out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUInt8(Advance: Boolean = True): UInt8;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt16(out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadInt16(Advance: Boolean = True): Int16;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt16(out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUInt16(Advance: Boolean = True): UInt16;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt32(out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadInt32(Advance: Boolean = True): Int32;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt32(out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUInt32(Advance: Boolean = True): UInt32;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt64(out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadInt64(Advance: Boolean = True): Int64;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt64(out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUInt64(Advance: Boolean = True): UInt64;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat32(out Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadFloat32(Advance: Boolean = True): Float32;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat64(out Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadFloat64(Advance: Boolean = True): Float64;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat80(out Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive10B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadFloat80(Advance: Boolean = True): Float80;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive10B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadDateTime(out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadDateTime(Advance: Boolean = True): TDateTime;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCurrency(out Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadCurrency(Advance: Boolean = True): Currency;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive8B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiChar(out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadAnsiChar(Advance: Boolean = True): AnsiChar;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8Char(out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUTF8Char(Advance: Boolean = True): UTF8Char;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive1B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideChar(out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadWideChar(Advance: Boolean = True): WideChar;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeChar(out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUnicodeChar(Advance: Boolean = True): UnicodeChar;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive2B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4Char(out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,SizeOf(Value),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUCS4Char(Advance: Boolean = True): UCS4Char;
begin
ReadValue(@Result,Advance,SizeOf(Result),vtPrimitive4B);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadChar(out Value: Char; Advance: Boolean = True): TMemSize;
var
  Temp: UInt16;
begin
Result := ReadValue(@Temp,Advance,SizeOf(Temp),vtPrimitive2B);
Value := Char(Temp);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadChar(Advance: Boolean = True): Char;
var
  Temp: UInt16;
begin
ReadValue(@Temp,Advance,SizeOf(Temp),vtPrimitive2B);
Result := Char(Temp);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadShortString(out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,0,vtShortString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadShortString(Advance: Boolean = True): ShortString;
begin
ReadValue(@Result,Advance,0,vtShortString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiString(out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,0,vtAnsiString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadAnsiString(Advance: Boolean = True): AnsiString;
begin
ReadValue(@Result,Advance,0,vtAnsiString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8String(out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,0,vtUTF8String);
end;
 
//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUTF8String(Advance: Boolean = True): UTF8String;
begin
ReadValue(@Result,Advance,0,vtUTF8String);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideString(out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,0,vtWideString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadWideString(Advance: Boolean = True): WideString;
begin
ReadValue(@Result,Advance,0,vtWideString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeString(out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,0,vtUnicodeString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUnicodeString(Advance: Boolean = True): UnicodeString;
begin
ReadValue(@Result,Advance,0,vtUnicodeString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4String(out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,0,vtUCS4String);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadUCS4String(Advance: Boolean = True): UCS4String;
begin
ReadValue(@Result,Advance,0,vtUCS4String);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadString(out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,0,vtString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadString(Advance: Boolean = True): String;
begin
ReadValue(@Result,Advance,0,vtString);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBuffer(var Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Buffer,Advance,Size,vtBytes);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadVariant(out Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(@Value,Advance,0,vtVariant);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LoadVariant(Advance: Boolean = True): Variant;
begin
ReadValue(@Result,Advance,0,vtVariant)
end;


{===============================================================================
--------------------------------------------------------------------------------
                                 TMemoryStreamer
--------------------------------------------------------------------------------
===============================================================================}
{===============================================================================
    TMemoryStreamer - class implementation
===============================================================================}
{-------------------------------------------------------------------------------
    TMemoryStreamer - protected methods
-------------------------------------------------------------------------------}

Function TMemoryStreamer.GetStartPtr: Pointer;
begin
{$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
Result := Pointer(PtrUInt(fStartPosition));
{$IFDEF FPCDWM}{$POP}{$ENDIF}
end;

//------------------------------------------------------------------------------

procedure TMemoryStreamer.SetBookmark(Index: Integer; Value: Int64);
begin
{
  casting to PtrInt is done to cut off higher 32bits on 32bit system
}
inherited SetBookmark(Index,Int64(PtrInt(Value)));
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetCurrentPosition: Int64;
begin
{$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
Result := Int64(PtrUInt(fCurrentPtr));
{$IFDEF FPCDWM}{$POP}{$ENDIF}
end;

//------------------------------------------------------------------------------

procedure TMemoryStreamer.SetCurrentPosition(NewPosition: Int64);
begin
{$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
fCurrentPtr := Pointer(PtrUInt(NewPosition));
{$IFDEF FPCDWM}{$POP}{$ENDIF}
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize;
begin
case ValueType of
  vtShortString:    Result := Ptr_WriteShortString(fCurrentPtr,ShortString(Value^),Advance,fEndian);
  vtAnsiString:     Result := Ptr_WriteAnsiString(fCurrentPtr,AnsiString(Value^),Advance,fEndian);
  vtUTF8String:     Result := Ptr_WriteUTF8String(fCurrentPtr,UTF8String(Value^),Advance,fEndian);
  vtWideString:     Result := Ptr_WriteWideString(fCurrentPtr,WideString(Value^),Advance,fEndian);
  vtUnicodeString:  Result := Ptr_WriteUnicodeString(fCurrentPtr,UnicodeString(Value^),Advance,fEndian);
  vtUCS4String:     Result := Ptr_WriteUCS4String(fCurrentPtr,UCS4String(Value^),Advance,fEndian);
  vtString:         Result := Ptr_WriteString(fCurrentPtr,String(Value^),Advance,fEndian);
  vtFillBytes:      Result := Ptr_FillBytes(fCurrentPtr,Size,UInt8(Value^),Advance,fEndian);
  vtPrimitive1B:    Result := Ptr_WriteUInt8(fCurrentPtr,UInt8(Value^),Advance,fEndian);
  vtPrimitive2B:    Result := Ptr_WriteUInt16(fCurrentPtr,UInt16(Value^),Advance,fEndian);
  vtPrimitive4B:    Result := Ptr_WriteUInt32(fCurrentPtr,UInt32(Value^),Advance,fEndian);
  vtPrimitive8B:    Result := Ptr_WriteUInt64(fCurrentPtr,UInt64(Value^),Advance,fEndian);
  vtPrimitive10B:   Result := Ptr_WriteFloat80(fCurrentPtr,Float80(Value^),Advance,fEndian);
  vtVariant:        Result := Ptr_WriteVariant(fCurrentPtr,Variant(Value^),Advance,fEndian);
else
 {vtBytes}
  Result := Ptr_WriteBuffer(fCurrentPtr,Value^,Size,Advance);
end;
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize;
begin
case ValueType of
  vtShortString:    Result := Ptr_ReadShortString(fCurrentPtr,ShortString(Value^),Advance,fEndian);
  vtAnsiString:     Result := Ptr_ReadAnsiString(fCurrentPtr,AnsiString(Value^),Advance,fEndian);
  vtUTF8String:     Result := Ptr_ReadUTF8String(fCurrentPtr,UTF8String(Value^),Advance,fEndian);
  vtWideString:     Result := Ptr_ReadWideString(fCurrentPtr,WideString(Value^),Advance,fEndian);
  vtUnicodeString:  Result := Ptr_ReadUnicodeString(fCurrentPtr,UnicodeString(Value^),Advance,fEndian);
  vtUCS4String:     Result := Ptr_ReadUCS4String(fCurrentPtr,UCS4String(Value^),Advance,fEndian);
  vtString:         Result := Ptr_ReadString(fCurrentPtr,String(Value^),Advance,fEndian);
  vtPrimitive1B:    Result := Ptr_ReadUInt8(fCurrentPtr,UInt8(Value^),Advance,fEndian);
  vtPrimitive2B:    Result := Ptr_ReadUInt16(fCurrentPtr,UInt16(Value^),Advance,fEndian);
  vtPrimitive4B:    Result := Ptr_ReadUInt32(fCurrentPtr,UInt32(Value^),Advance,fEndian);
  vtPrimitive8B:    Result := Ptr_ReadUInt64(fCurrentPtr,UInt64(Value^),Advance,fEndian);
  vtPrimitive10B:   Result := Ptr_ReadFloat80(fCurrentPtr,Float80(Value^),Advance,fEndian);
  vtVariant:        Result := Ptr_ReadVariant(fCurrentPtr,Variant(Value^),Advance,fEndian);
else
 {vtFillBytes, vtBytes}
  Result := Ptr_ReadBuffer(fCurrentPtr,Value^,Size,Advance);
end;
end;

//------------------------------------------------------------------------------

procedure TMemoryStreamer.Initialize(Memory: Pointer);
begin
inherited Initialize;
fCurrentPtr := Memory;
{$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
fStartPosition := Int64(PtrUInt(Memory));
{$IFDEF FPCDWM}{$POP}{$ENDIF}
end;

{-------------------------------------------------------------------------------
    TMemoryStreamer - public methods
-------------------------------------------------------------------------------}

constructor TMemoryStreamer.Create(Memory: Pointer);
begin
inherited Create;
Initialize(Memory);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.IndexOfBookmark(Position: Int64): Integer;
begin
Result := inherited IndexOfBookmark(PtrInt(Position));
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.AddBookmark(Position: Int64): Integer;
begin
Result := inherited AddBookmark(PtrInt(Position));
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.RemoveBookmark(Position: Int64; RemoveAll: Boolean = True): Integer;
begin
Result := inherited RemoveBookmark(PtrInt(Position),RemoveAll);
end;


{===============================================================================
--------------------------------------------------------------------------------
                                 TStreamStreamer
--------------------------------------------------------------------------------
===============================================================================}
{===============================================================================
    TStreamStreamer - class implementation
===============================================================================}
{-------------------------------------------------------------------------------
    TStreamStreamer - protected methods
-------------------------------------------------------------------------------}

Function TStreamStreamer.GetCurrentPosition: Int64;
begin
Result := fTarget.Position;
end;

//------------------------------------------------------------------------------

procedure TStreamStreamer.SetCurrentPosition(NewPosition: Int64);
begin
fTarget.Position := NewPosition;
end;

//------------------------------------------------------------------------------

Function TStreamStreamer.WriteValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize;
begin
case ValueType of
  vtShortString:    Result := Stream_WriteShortString(fTarget,ShortString(Value^),Advance);
  vtAnsiString:     Result := Stream_WriteAnsiString(fTarget,AnsiString(Value^),Advance);
  vtUTF8String:     Result := Stream_WriteUTF8String(fTarget,UTF8String(Value^),Advance);
  vtWideString:     Result := Stream_WriteWideString(fTarget,WideString(Value^),Advance);
  vtUnicodeString:  Result := Stream_WriteUnicodeString(fTarget,UnicodeString(Value^),Advance);
  vtUCS4String:     Result := Stream_WriteUCS4String(fTarget,UCS4String(Value^),Advance);
  vtString:         Result := Stream_WriteString(fTarget,String(Value^),Advance);
  vtFillBytes:      Result := Stream_FillBytes(fTarget,Size,UInt8(Value^),Advance);
  vtPrimitive1B:    Result := Stream_WriteUInt8(fTarget,UInt8(Value^),Advance);
  vtPrimitive2B:    Result := Stream_WriteUInt16(fTarget,UInt16(Value^),Advance);
  vtPrimitive4B:    Result := Stream_WriteUInt32(fTarget,UInt32(Value^),Advance);
  vtPrimitive8B:    Result := Stream_WriteUInt64(fTarget,UInt64(Value^),Advance);
  vtPrimitive10B:   Result := Stream_WriteFloat80(fTarget,Float80(Value^),Advance);
  vtVariant:        Result := Stream_WriteVariant(fTarget,Variant(Value^),Advance);
else
 {vtBytes}
  Result := Stream_WriteBuffer(fTarget,Value^,Size,Advance);
end;
end;

//------------------------------------------------------------------------------

Function TStreamStreamer.ReadValue(Value: Pointer; Advance: Boolean; Size: TMemSize; ValueType: TValueType): TMemSize;
begin
case ValueType of
  vtShortString:    Result := Stream_ReadShortString(fTarget,ShortString(Value^),Advance);
  vtAnsiString:     Result := Stream_ReadAnsiString(fTarget,AnsiString(Value^),Advance);
  vtUTF8String:     Result := Stream_ReadUTF8String(fTarget,UTF8String(Value^),Advance);
  vtWideString:     Result := Stream_ReadWideString(fTarget,WideString(Value^),Advance);
  vtUnicodeString:  Result := Stream_ReadUnicodeString(fTarget,UnicodeString(Value^),Advance);
  vtUCS4String:     Result := Stream_ReadUCS4String(fTarget,UCS4String(Value^),Advance);
  vtString:         Result := Stream_ReadString(fTarget,String(Value^),Advance);
  vtPrimitive1B:    Result := Stream_ReadUInt8(fTarget,UInt8(Value^),Advance);
  vtPrimitive2B:    Result := Stream_ReadUInt16(fTarget,UInt16(Value^),Advance);
  vtPrimitive4B:    Result := Stream_ReadUInt32(fTarget,UInt32(Value^),Advance);
  vtPrimitive8B:    Result := Stream_ReadUInt64(fTarget,UInt64(Value^),Advance);
  vtPrimitive10B:   Result := Stream_ReadFloat80(fTarget,Float80(Value^),Advance);
  vtVariant:        Result := Stream_ReadVariant(fTarget,Variant(Value^),Advance);
else
 {vtFillBytes,vtBytes}
  Result := Stream_ReadBuffer(fTarget,Value^,Size,Advance);
end;
end;

//------------------------------------------------------------------------------

procedure TStreamStreamer.Initialize(Target: TStream);
begin
inherited Initialize;
fTarget := Target;
fStartPosition := Target.Position;
end;

{-------------------------------------------------------------------------------
    TStreamStreamer - public methods
-------------------------------------------------------------------------------}

constructor TStreamStreamer.Create(Target: TStream);
begin
inherited Create;
Initialize(Target);
end;
*)

{===============================================================================
--------------------------------------------------------------------------------
                                Unit preparation
--------------------------------------------------------------------------------
===============================================================================}

procedure UnitInitialize;

  Function GetBOAStride(const Arr: array of UInt8): Integer;
  begin
  {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
    Result := Integer(PtrUInt(Addr(Arr[1])) - PtrUInt(Addr(Arr[0])));
  {$IFDEF FPCDWM}{$POP}{$ENDIF}
  end;

begin
ByteOpenArrayIsPacked := GetBOAStride([0,0]) = 1;
end;

//==============================================================================

initialization
  UnitInitialize;

end.

