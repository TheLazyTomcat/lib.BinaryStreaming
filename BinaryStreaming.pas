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
    including in UTF-16 strings, ...) can be written either with little endian
    (LE) or big endian (BE) byte order. For single-byte primitives, the byte
    order can also be selected, but it has no effect on the result (a common
    code is called).

      WARNING - the endianness is not stored with the data, it is up to you
                to use proper loading function for the byte order used during
                saving.  

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

    Default type String is always stored as an UTF-8 encoded string, again
    irrespective of how it is declared.

    Buffers and array of bytes are both stored as plain byte streams, without
    explicit size.

    Variants are stored in a litle more complex way - first a byte denoting
    type of the variant is stored (note that value of this byte does NOT
    correspond to a TVarType value), directly followed by the value itself.
    The value is stored as if it was a normal variable (eg. for varWord variant
    the function WriteUInt16 is called).

    For variant arrays, immediately after the type byte a dimension count is
    stored (int32), followed by an array of indices bounds (two int32 values
    for each dimension, first low bound followed by a high bound, ordered from
    lowest dimension to highest dimension). After this, an array of items is
    stored, each item is stored as a separate variant. When saving the items,
    the right-most index is incremented first.

      NOTE - only standard variant types are supported, and from them only the
             "streamable" types (boolean, integers, floats, strings) and their
             arrays are allowed - so no empty/null variants, interfaces,
             objects, errors, and so on.

    Also, since some older compilers do not support all in-here used variant
    types (namely varUInt64 and varUString), there are some specific rules in
    effect when saving and loading such types in programs compiled by those
    compilers...

      Localy unsupported types are not saved at all, simply because a variant
      of that type cannot be even created (note that the type UInt64 is
      sometimes declared only an alias for Int64, so when creating variant from
      that type, it will be of type varInt64, in that case such variant will be
      saved as varInt64).

      When loading unsupported varUInt64, it will be loaded as Int64 (with the
      same BINARY data, ie. not necessarily the same numerical value).

      When loading unsupported varUString (unicode string), it is normally
      loaded, but the returned variant will be of type varOleStr (wide string).

    Parameter Advance in writing and reading functions indicates whether the
    position in stream being written to or read from (or the passed memory
    pointer) can be advanced by number of bytes written or read. When set to
    true, the position (pointer) is advanced, when false, the position is the
    same after the call as was before it.

    Return value of all read and write functions is number of bytes written or
    read (does not apply to Get* functions, as they are returning the value
    being read).

  Version 2.0.3 (2024-04-14)

  Last change 2024-04-14

  ©2015-2024 František Milt

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
    AuxClasses    - github.com/TheLazyTomcat/Lib.AuxClasses
  * AuxExceptions - github.com/TheLazyTomcat/Lib.AuxExceptions
    AuxTypes      - github.com/TheLazyTomcat/Lib.AuxTypes
    StrRect       - github.com/TheLazyTomcat/Lib.StrRect

  Library AuxExceptions is required only when rebasing local exception classes
  (see symbol BinaryStreaming_UseAuxExceptions for details).

  Library AuxExceptions might also be required as an indirect dependency.

  Indirect dependencies:
    SimpleCPUID - github.com/TheLazyTomcat/Lib.SimpleCPUID
    UInt64Utils - github.com/TheLazyTomcat/Lib.UInt64Utils
    WinFileInfo - github.com/TheLazyTomcat/Lib.WinFileInfo

===============================================================================}
unit BinaryStreaming;
(*
  All comments in the form {<lite-...>} are tags used to compress this unit
  into BinaryStreamingLite. They have no impact on this library and should be
  ignored.
*)
{<lite-unit>}
{
  BinaryStreaming_UseAuxExceptions

  If you want library-specific exceptions to be based on more advanced classes
  provided by AuxExceptions library instead of basic Exception class, and don't
  want to or cannot change code in this unit, you can define global symbol
  BinaryStreaming_UseAuxExceptions to achieve this.
}
{$IF Defined(BinaryStreaming_UseAuxExceptions)}
  {$DEFINE UseAuxExceptions}
{$IFEND}

//------------------------------------------------------------------------------

{<lite-begin>}
{<lite-ln>}
{$IFDEF FPC}
  {$MODE ObjFPC}
  {<lite-remove>}
  {$MODESWITCH DuplicateLocals+}
  {$INLINE ON}
  {$DEFINE CanInline}
  {$DEFINE FPC_DisableWarns}
  {$MACRO ON}
{$ELSE}
  {$IF CompilerVersion >= 17} // Delphi 2005+
    {$DEFINE CanInline}
  {$IFEND}
{$ENDIF}
{$H+}
{<lite-end-ln>}

//------------------------------------------------------------------------------
// do not change following defines

{$UNDEF BS_OverflowChecks}

{$UNDEF BS_INC_VW}  // variant writing
{$UNDEF BS_INC_VR}  // variant reading
{$UNDEF BS_INC_M}   // operating on memory (stream is inferred when not defined)
{$UNDEF BS_INC_L}   // little endian (big endian inferred when not defined)

interface

{<lite-interface>}

uses
  SysUtils, Classes,
  AuxTypes, AuxClasses{$IFNDEF FPC}, StrRect{$ENDIF}
  {$IFDEF UseAuxExceptions}, AuxExceptions{$ENDIF};

{<lite-begin>}
{<lite-block-replace-begin EBS EBSL>}
{===============================================================================
    Library-specific exceptions
===============================================================================}
type
  EBSException = class({$IFDEF UseAuxExceptions}EAEGeneralException{$ELSE}Exception{$ENDIF});
{<lite-end-ln>}

  EBSUnsupportedVarType = class(EBSException);
  EBSIndexOutOfBounds   = class(EBSException);
  EBSInvalidValue       = class(EBSException);
  EBSDuplicateItem      = class(EBSException);
  EBSUnknownItem        = class(EBSException);

  EBSStringTooLong = class(EBSInvalidValue);

{===============================================================================
    Endianness - declaration
===============================================================================}
type
  TEndian = (endLittle,endBig,endSystem,endDefault{endLittle});

const
  SysEndian = {$IFDEF ENDIAN_BIG}endBig{$ELSE}endLittle{$ENDIF};

{
  ProbeEndian

  Performs runtime test of endianness - it returns the actual true byte order,
  not whatever was assigned to SysEndian during compilation (though these two
  values should normally match).

  It can only return endLittle or endBig, if the probing shows anything else
  (eg. mixed endianness or differing endianness depending on a variable type),
  then the function will raise an EBSInvalidValue exception.

  The byte order is tested on a 4-byte (32bit) signed integer and 4-byte float.
}
Function ProbeEndian: TEndian;

{
  ResolveEndian

  Resolves endSystem to a value of constant SysEndian and endDefault to
  endLittle, other values are returned unchanged.
}
Function ResolveEndian(Endian: TEndian): TEndian;

{<lite-begin>}
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

Function StreamedSize_ShortString(const Value: ShortString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_AnsiString(const Value: AnsiString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UTF8String(const Value: UTF8String): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_WideString(const Value: WideString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UnicodeString(const Value: UnicodeString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_UCS4String(const Value: UCS4String): TMemSize;
Function StreamedSize_String(const Value: String): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------
{<lite-end-ln>}
Function StreamedSize_SmallAnsiString(const Value: AnsiString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_SmallUTF8String(const Value: UTF8String): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_SmallWideString(const Value: WideString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_SmallUnicodeString(const Value: UnicodeString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_SmallUCS4String(const Value: UCS4String): TMemSize;
Function StreamedSize_SmallString(const Value: String): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------

Function StreamedSize_TinyAnsiString(const Value: AnsiString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_TinyUTF8String(const Value: UTF8String): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_TinyWideString(const Value: WideString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_TinyUnicodeString(const Value: UnicodeString): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_TinyUCS4String(const Value: UCS4String): TMemSize;
Function StreamedSize_TinyString(const Value: String): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

//------------------------------------------------------------------------------
{<lite-begin>}
Function StreamedSize_Buffer(Size: TMemSize): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
Function StreamedSize_Bytes(Count: TMemSize): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function StreamedSize_Variant(const Value: Variant): TMemSize;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{===============================================================================
--------------------------------------------------------------------------------
                                 Memory writing
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function Ptr_WriteBool_LE(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_WriteBool_LE(Dest: Pointer; Value: ByteBool): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_WriteBool_BE(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBool_BE(Dest: Pointer; Value: ByteBool): TMemSize; overload;

Function Ptr_WriteBool(var Dest: Pointer; Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteBool(Dest: Pointer; Value: ByteBool; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteBoolean_LE(var Dest: Pointer; Value: Boolean; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteBoolean_LE(Dest: Pointer; Value: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteBoolean_BE(var Dest: Pointer; Value: Boolean; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteBoolean_BE(Dest: Pointer; Value: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteBoolean(var Dest: Pointer; Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteBoolean(Dest: Pointer; Value: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function Ptr_WriteInt8_LE(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_WriteInt8_LE(Dest: Pointer; Value: Int8): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_WriteInt8_BE(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt8_BE(Dest: Pointer; Value: Int8): TMemSize; overload;

Function Ptr_WriteInt8(var Dest: Pointer; Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteInt8(Dest: Pointer; Value: Int8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUInt8_LE(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_WriteUInt8_LE(Dest: Pointer; Value: UInt8): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_WriteUInt8_BE(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt8_BE(Dest: Pointer; Value: UInt8): TMemSize; overload;

Function Ptr_WriteUInt8(var Dest: Pointer; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUInt8(Dest: Pointer; Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteInt16_LE(var Dest: Pointer; Value: Int16; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt16_LE(Dest: Pointer; Value: Int16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteInt16_BE(var Dest: Pointer; Value: Int16; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt16_BE(Dest: Pointer; Value: Int16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt16(var Dest: Pointer; Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteInt16(Dest: Pointer; Value: Int16; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUInt16_LE(var Dest: Pointer; Value: UInt16; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt16_LE(Dest: Pointer; Value: UInt16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteUInt16_BE(var Dest: Pointer; Value: UInt16; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt16_BE(Dest: Pointer; Value: UInt16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt16(var Dest: Pointer; Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUInt16(Dest: Pointer; Value: UInt16; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteInt32_LE(var Dest: Pointer; Value: Int32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt32_LE(Dest: Pointer; Value: Int32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteInt32_BE(var Dest: Pointer; Value: Int32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt32_BE(Dest: Pointer; Value: Int32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt32(var Dest: Pointer; Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteInt32(Dest: Pointer; Value: Int32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUInt32_LE(var Dest: Pointer; Value: UInt32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt32_LE(Dest: Pointer; Value: UInt32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteUInt32_BE(var Dest: Pointer; Value: UInt32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt32_BE(Dest: Pointer; Value: UInt32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt32(var Dest: Pointer; Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUInt32(Dest: Pointer; Value: UInt32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteInt64_LE(var Dest: Pointer; Value: Int64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt64_LE(Dest: Pointer; Value: Int64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteInt64_BE(var Dest: Pointer; Value: Int64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteInt64_BE(Dest: Pointer; Value: Int64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteInt64(var Dest: Pointer; Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteInt64(Dest: Pointer; Value: Int64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUInt64_LE(var Dest: Pointer; Value: UInt64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt64_LE(Dest: Pointer; Value: UInt64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteUInt64_BE(var Dest: Pointer; Value: UInt64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUInt64_BE(Dest: Pointer; Value: UInt64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUInt64(var Dest: Pointer; Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUInt64(Dest: Pointer; Value: UInt64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Ptr_WriteFloat32_LE(var Dest: Pointer; Value: Float32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat32_LE(Dest: Pointer; Value: Float32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteFloat32_BE(var Dest: Pointer; Value: Float32; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat32_BE(Dest: Pointer; Value: Float32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat32(var Dest: Pointer; Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteFloat32(Dest: Pointer; Value: Float32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">} 
//------------------------------------------------------------------------------

Function Ptr_WriteFloat64_LE(var Dest: Pointer; Value: Float64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat64_LE(Dest: Pointer; Value: Float64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteFloat64_BE(var Dest: Pointer; Value: Float64; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat64_BE(Dest: Pointer; Value: Float64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat64(var Dest: Pointer; Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteFloat64(Dest: Pointer; Value: Float64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteFloat80_LE(var Dest: Pointer; Value: Float80; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat80_LE(Dest: Pointer; Value: Float80): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteFloat80_BE(var Dest: Pointer; Value: Float80; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteFloat80_BE(Dest: Pointer; Value: Float80): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteFloat80(var Dest: Pointer; Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteFloat80(Dest: Pointer; Value: Float80; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteDateTime_LE(var Dest: Pointer; Value: TDateTime; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteDateTime_LE(Dest: Pointer; Value: TDateTime): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteDateTime_BE(var Dest: Pointer; Value: TDateTime; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteDateTime_BE(Dest: Pointer; Value: TDateTime): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteDateTime(var Dest: Pointer; Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteDateTime(Dest: Pointer; Value: TDateTime; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteCurrency_LE(var Dest: Pointer; Value: Currency; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteCurrency_LE(Dest: Pointer; Value: Currency): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteCurrency_BE(var Dest: Pointer; Value: Currency; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteCurrency_BE(Dest: Pointer; Value: Currency): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteCurrency(var Dest: Pointer; Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteCurrency(Dest: Pointer; Value: Currency; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function Ptr_WriteAnsiChar_LE(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_WriteAnsiChar_LE(Dest: Pointer; Value: AnsiChar): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_WriteAnsiChar_BE(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteAnsiChar_BE(Dest: Pointer; Value: AnsiChar): TMemSize; overload;

Function Ptr_WriteAnsiChar(var Dest: Pointer; Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteAnsiChar(Dest: Pointer; Value: AnsiChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUTF8Char_LE(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_WriteUTF8Char_LE(Dest: Pointer; Value: UTF8Char): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_WriteUTF8Char_BE(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUTF8Char_BE(Dest: Pointer; Value: UTF8Char): TMemSize; overload;

Function Ptr_WriteUTF8Char(var Dest: Pointer; Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUTF8Char(Dest: Pointer; Value: UTF8Char; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteWideChar_LE(var Dest: Pointer; Value: WideChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteWideChar_LE(Dest: Pointer; Value: WideChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteWideChar_BE(var Dest: Pointer; Value: WideChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteWideChar_BE(Dest: Pointer; Value: WideChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteWideChar(var Dest: Pointer; Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteWideChar(Dest: Pointer; Value: WideChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeChar_LE(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUnicodeChar_LE(Dest: Pointer; Value: UnicodeChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteUnicodeChar_BE(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUnicodeChar_BE(Dest: Pointer; Value: UnicodeChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUnicodeChar(var Dest: Pointer; Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUnicodeChar(Dest: Pointer; Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUCS4Char_LE(var Dest: Pointer; Value: UCS4Char; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUCS4Char_LE(Dest: Pointer; Value: UCS4Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteUCS4Char_BE(var Dest: Pointer; Value: UCS4Char; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUCS4Char_BE(Dest: Pointer; Value: UCS4Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUCS4Char(var Dest: Pointer; Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUCS4Char(Dest: Pointer; Value: UCS4Char; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteChar_LE(var Dest: Pointer; Value: Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteChar_LE(Dest: Pointer; Value: Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteChar_BE(var Dest: Pointer; Value: Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteChar_BE(Dest: Pointer; Value: Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteChar(var Dest: Pointer; Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteChar(Dest: Pointer; Value: Char; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function Ptr_WriteShortString_LE(var Dest: Pointer; const Value: ShortString; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_WriteShortString_LE(Dest: Pointer; const Value: ShortString): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_WriteShortString_BE(var Dest: Pointer; const Value: ShortString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteShortString_BE(Dest: Pointer; const Value: ShortString): TMemSize; overload;

Function Ptr_WriteShortString(var Dest: Pointer; const Value: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteShortString(Dest: Pointer; const Value: ShortString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteAnsiString_LE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteAnsiString_LE(Dest: Pointer; const Value: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteAnsiString_BE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteAnsiString_BE(Dest: Pointer; const Value: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteAnsiString(var Dest: Pointer; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteAnsiString(Dest: Pointer; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUTF8String_LE(var Dest: Pointer; const Value: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUTF8String_LE(Dest: Pointer; const Value: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteUTF8String_BE(var Dest: Pointer; const Value: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUTF8String_BE(Dest: Pointer; const Value: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUTF8String(var Dest: Pointer; const Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUTF8String(Dest: Pointer; const Value: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteWideString_LE(var Dest: Pointer; const Value: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteWideString_LE(Dest: Pointer; const Value: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteWideString_BE(var Dest: Pointer; const Value: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteWideString_BE(Dest: Pointer; const Value: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteWideString(var Dest: Pointer; const Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteWideString(Dest: Pointer; const Value: WideString; Endian: TEndian = endDefault): TMemSize; overload;
 
{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeString_LE(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUnicodeString_LE(Dest: Pointer; const Value: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteUnicodeString_BE(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUnicodeString_BE(Dest: Pointer; const Value: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUnicodeString(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUnicodeString(Dest: Pointer; const Value: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteUCS4String_LE(var Dest: Pointer; const Value: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUCS4String_LE(Dest: Pointer; const Value: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteUCS4String_BE(var Dest: Pointer; const Value: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteUCS4String_BE(Dest: Pointer; const Value: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteUCS4String(var Dest: Pointer; const Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteUCS4String(Dest: Pointer; const Value: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_WriteString_LE(var Dest: Pointer; const Value: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteString_LE(Dest: Pointer; const Value: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_WriteString_BE(var Dest: Pointer; const Value: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteString_BE(Dest: Pointer; const Value: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteString(var Dest: Pointer; const Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteString(Dest: Pointer; const Value: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Small strings
-------------------------------------------------------------------------------}

Function Ptr_WriteSmallAnsiString_LE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallAnsiString_LE(Dest: Pointer; const Value: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallAnsiString_BE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallAnsiString_BE(Dest: Pointer; const Value: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallAnsiString(var Dest: Pointer; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteSmallAnsiString(Dest: Pointer; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUTF8String_LE(var Dest: Pointer; const Value: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallUTF8String_LE(Dest: Pointer; const Value: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallUTF8String_BE(var Dest: Pointer; const Value: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallUTF8String_BE(Dest: Pointer; const Value: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallUTF8String(var Dest: Pointer; const Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteSmallUTF8String(Dest: Pointer; const Value: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallWideString_LE(var Dest: Pointer; const Value: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallWideString_LE(Dest: Pointer; const Value: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallWideString_BE(var Dest: Pointer; const Value: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallWideString_BE(Dest: Pointer; const Value: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallWideString(var Dest: Pointer; const Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteSmallWideString(Dest: Pointer; const Value: WideString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUnicodeString_LE(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallUnicodeString_LE(Dest: Pointer; const Value: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallUnicodeString_BE(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallUnicodeString_BE(Dest: Pointer; const Value: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallUnicodeString(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteSmallUnicodeString(Dest: Pointer; const Value: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUCS4String_LE(var Dest: Pointer; const Value: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallUCS4String_LE(Dest: Pointer; const Value: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallUCS4String_BE(var Dest: Pointer; const Value: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteSmallUCS4String_BE(Dest: Pointer; const Value: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallUCS4String(var Dest: Pointer; const Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteSmallUCS4String(Dest: Pointer; const Value: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallString_LE(var Dest: Pointer; const Value: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteSmallString_LE(Dest: Pointer; const Value: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallString_BE(var Dest: Pointer; const Value: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteSmallString_BE(Dest: Pointer; const Value: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_WriteSmallString(var Dest: Pointer; const Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_WriteSmallString(Dest: Pointer; const Value: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Tiny strings
-------------------------------------------------------------------------------}

Function Ptr_WriteTinyAnsiString_LE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteTinyAnsiString_LE(Dest: Pointer; const Value: AnsiString): TMemSize; overload;

Function Ptr_WriteTinyAnsiString_BE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteTinyAnsiString_BE(Dest: Pointer; const Value: AnsiString): TMemSize; overload;

Function Ptr_WriteTinyAnsiString(var Dest: Pointer; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteTinyAnsiString(Dest: Pointer; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function Ptr_WriteBuffer_LE(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBuffer_LE(Dest: Pointer; const Buffer; Size: TMemSize): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_WriteBuffer_BE(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBuffer_BE(Dest: Pointer; const Buffer; Size: TMemSize): TMemSize; overload;

Function Ptr_WriteBuffer(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteBuffer(Dest: Pointer; const Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------
// note - calls with open array cannot be inlined both in Delphi and FPC

Function Ptr_WriteBytes_LE(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBytes_LE(Dest: Pointer; const Value: array of UInt8): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_WriteBytes_BE(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize; overload;
Function Ptr_WriteBytes_BE(Dest: Pointer; const Value: array of UInt8): TMemSize; overload;

Function Ptr_WriteBytes(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_WriteBytes(Dest: Pointer; const Value: array of UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_FillBytes_LE(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_FillBytes_LE(Dest: Pointer; Count: TMemSize; Value: UInt8): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_FillBytes_BE(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize; overload;
Function Ptr_FillBytes_BE(Dest: Pointer; Count: TMemSize; Value: UInt8): TMemSize; overload;

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


{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{===============================================================================
--------------------------------------------------------------------------------
                                 Memory reading
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

Function Ptr_ReadBool_LE(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_ReadBool_LE(Src: Pointer; out Value: ByteBool): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_ReadBool_BE(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadBool_BE(Src: Pointer; out Value: ByteBool): TMemSize; overload;

Function Ptr_ReadBool(var Src: Pointer; out Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadBool(Src: Pointer; out Value: ByteBool; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetBool_LE(var Src: Pointer; Advance: Boolean): ByteBool; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetBool_LE(Src: Pointer): ByteBool; overload;
{<lite-end-ln>}

Function Ptr_GetBool_BE(var Src: Pointer; Advance: Boolean): ByteBool; overload;
Function Ptr_GetBool_BE(Src: Pointer): ByteBool; overload;

Function Ptr_GetBool(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): ByteBool; overload;
Function Ptr_GetBool(Src: Pointer; Endian: TEndian = endDefault): ByteBool; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadBoolean_LE(var Src: Pointer; out Value: Boolean; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadBoolean_LE(Src: Pointer; out Value: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadBoolean_BE(var Src: Pointer; out Value: Boolean; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadBoolean_BE(Src: Pointer; out Value: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadBoolean(var Src: Pointer; out Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadBoolean(Src: Pointer; out Value: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetBoolean_LE(var Src: Pointer; Advance: Boolean): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetBoolean_LE(Src: Pointer): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetBoolean_BE(var Src: Pointer; Advance: Boolean): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetBoolean_BE(Src: Pointer): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetBoolean(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetBoolean(Src: Pointer; Endian: TEndian = endDefault): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

Function Ptr_ReadInt8_LE(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_ReadInt8_LE(Src: Pointer; out Value: Int8): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_ReadInt8_BE(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt8_BE(Src: Pointer; out Value: Int8): TMemSize; overload;

Function Ptr_ReadInt8(var Src: Pointer; out Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadInt8(Src: Pointer; out Value: Int8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetInt8_LE(var Src: Pointer; Advance: Boolean): Int8; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetInt8_LE(Src: Pointer): Int8; overload;
{<lite-end-ln>}

Function Ptr_GetInt8_BE(var Src: Pointer; Advance: Boolean): Int8; overload;
Function Ptr_GetInt8_BE(Src: Pointer): Int8; overload;

Function Ptr_GetInt8(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int8; overload;
Function Ptr_GetInt8(Src: Pointer; Endian: TEndian = endDefault): Int8; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUInt8_LE(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_ReadUInt8_LE(Src: Pointer; out Value: UInt8): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_ReadUInt8_BE(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt8_BE(Src: Pointer; out Value: UInt8): TMemSize; overload;

Function Ptr_ReadUInt8(var Src: Pointer; out Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUInt8(Src: Pointer; out Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetUInt8_LE(var Src: Pointer; Advance: Boolean): UInt8; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetUInt8_LE(Src: Pointer): UInt8; overload;
{<lite-end-ln>}

Function Ptr_GetUInt8_BE(var Src: Pointer; Advance: Boolean): UInt8; overload;
Function Ptr_GetUInt8_BE(Src: Pointer): UInt8; overload;

Function Ptr_GetUInt8(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt8; overload;
Function Ptr_GetUInt8(Src: Pointer; Endian: TEndian = endDefault): UInt8; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadInt16_LE(var Src: Pointer; out Value: Int16; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt16_LE(Src: Pointer; out Value: Int16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadInt16_BE(var Src: Pointer; out Value: Int16; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt16_BE(Src: Pointer; out Value: Int16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt16(var Src: Pointer; out Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadInt16(Src: Pointer; out Value: Int16; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetInt16_LE(var Src: Pointer; Advance: Boolean): Int16; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt16_LE(Src: Pointer): Int16; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetInt16_BE(var Src: Pointer; Advance: Boolean): Int16; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt16_BE(Src: Pointer): Int16; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt16(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int16; overload;
Function Ptr_GetInt16(Src: Pointer; Endian: TEndian = endDefault): Int16; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUInt16_LE(var Src: Pointer; out Value: UInt16; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt16_LE(Src: Pointer; out Value: UInt16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadUInt16_BE(var Src: Pointer; out Value: UInt16; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt16_BE(Src: Pointer; out Value: UInt16): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt16(var Src: Pointer; out Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUInt16(Src: Pointer; out Value: UInt16; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetUInt16_LE(var Src: Pointer; Advance: Boolean): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt16_LE(Src: Pointer): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetUInt16_BE(var Src: Pointer; Advance: Boolean): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt16_BE(Src: Pointer): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt16(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt16; overload;
Function Ptr_GetUInt16(Src: Pointer; Endian: TEndian = endDefault): UInt16; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadInt32_LE(var Src: Pointer; out Value: Int32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt32_LE(Src: Pointer; out Value: Int32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadInt32_BE(var Src: Pointer; out Value: Int32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt32_BE(Src: Pointer; out Value: Int32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt32(var Src: Pointer; out Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadInt32(Src: Pointer; out Value: Int32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetInt32_LE(var Src: Pointer; Advance: Boolean): Int32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt32_LE(Src: Pointer): Int32; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetInt32_BE(var Src: Pointer; Advance: Boolean): Int32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt32_BE(Src: Pointer): Int32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int32; overload;
Function Ptr_GetInt32(Src: Pointer; Endian: TEndian = endDefault): Int32; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUInt32_LE(var Src: Pointer; out Value: UInt32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt32_LE(Src: Pointer; out Value: UInt32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadUInt32_BE(var Src: Pointer; out Value: UInt32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt32_BE(Src: Pointer; out Value: UInt32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt32(var Src: Pointer; out Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUInt32(Src: Pointer; out Value: UInt32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetUInt32_LE(var Src: Pointer; Advance: Boolean): UInt32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt32_LE(Src: Pointer): UInt32; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetUInt32_BE(var Src: Pointer; Advance: Boolean): UInt32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt32_BE(Src: Pointer): UInt32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt32; overload;
Function Ptr_GetUInt32(Src: Pointer; Endian: TEndian = endDefault): UInt32; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadInt64_LE(var Src: Pointer; out Value: Int64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt64_LE(Src: Pointer; out Value: Int64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadInt64_BE(var Src: Pointer; out Value: Int64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadInt64_BE(Src: Pointer; out Value: Int64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadInt64(var Src: Pointer; out Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadInt64(Src: Pointer; out Value: Int64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetInt64_LE(var Src: Pointer; Advance: Boolean): Int64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt64_LE(Src: Pointer): Int64; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetInt64_BE(var Src: Pointer; Advance: Boolean): Int64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetInt64_BE(Src: Pointer): Int64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetInt64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Int64; overload;
Function Ptr_GetInt64(Src: Pointer; Endian: TEndian = endDefault): Int64; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUInt64_LE(var Src: Pointer; out Value: UInt64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt64_LE(Src: Pointer; out Value: UInt64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadUInt64_BE(var Src: Pointer; out Value: UInt64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUInt64_BE(Src: Pointer; out Value: UInt64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUInt64(var Src: Pointer; out Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUInt64(Src: Pointer; out Value: UInt64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetUInt64_LE(var Src: Pointer; Advance: Boolean): UInt64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt64_LE(Src: Pointer): UInt64; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetUInt64_BE(var Src: Pointer; Advance: Boolean): UInt64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUInt64_BE(Src: Pointer): UInt64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUInt64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UInt64; overload;
Function Ptr_GetUInt64(Src: Pointer; Endian: TEndian = endDefault): UInt64; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

Function Ptr_ReadFloat32_LE(var Src: Pointer; out Value: Float32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat32_LE(Src: Pointer; out Value: Float32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadFloat32_BE(var Src: Pointer; out Value: Float32; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat32_BE(Src: Pointer; out Value: Float32): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat32(var Src: Pointer; out Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadFloat32(Src: Pointer; out Value: Float32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetFloat32_LE(var Src: Pointer; Advance: Boolean): Float32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat32_LE(Src: Pointer): Float32; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetFloat32_BE(var Src: Pointer; Advance: Boolean): Float32; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat32_BE(Src: Pointer): Float32; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat32(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float32; overload;
Function Ptr_GetFloat32(Src: Pointer; Endian: TEndian = endDefault): Float32; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadFloat64_LE(var Src: Pointer; out Value: Float64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat64_LE(Src: Pointer; out Value: Float64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadFloat64_BE(var Src: Pointer; out Value: Float64; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat64_BE(Src: Pointer; out Value: Float64): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat64(var Src: Pointer; out Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadFloat64(Src: Pointer; out Value: Float64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetFloat64_LE(var Src: Pointer; Advance: Boolean): Float64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat64_LE(Src: Pointer): Float64; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetFloat64_BE(var Src: Pointer; Advance: Boolean): Float64; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat64_BE(Src: Pointer): Float64; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat64(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float64; overload;
Function Ptr_GetFloat64(Src: Pointer; Endian: TEndian = endDefault): Float64; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadFloat80_LE(var Src: Pointer; out Value: Float80; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat80_LE(Src: Pointer; out Value: Float80): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadFloat80_BE(var Src: Pointer; out Value: Float80; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadFloat80_BE(Src: Pointer; out Value: Float80): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadFloat80(var Src: Pointer; out Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadFloat80(Src: Pointer; out Value: Float80; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetFloat80_LE(var Src: Pointer; Advance: Boolean): Float80; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat80_LE(Src: Pointer): Float80; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetFloat80_BE(var Src: Pointer; Advance: Boolean): Float80; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetFloat80_BE(Src: Pointer): Float80; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetFloat80(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Float80; overload;
Function Ptr_GetFloat80(Src: Pointer; Endian: TEndian = endDefault): Float80; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadDateTime_LE(var Src: Pointer; out Value: TDateTime; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadDateTime_LE(Src: Pointer; out Value: TDateTime): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadDateTime_BE(var Src: Pointer; out Value: TDateTime; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadDateTime_BE(Src: Pointer; out Value: TDateTime): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadDateTime(var Src: Pointer; out Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadDateTime(Src: Pointer; out Value: TDateTime; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetDateTime_LE(var Src: Pointer; Advance: Boolean): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetDateTime_LE(Src: Pointer): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetDateTime_BE(var Src: Pointer; Advance: Boolean): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetDateTime_BE(Src: Pointer): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetDateTime(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetDateTime(Src: Pointer; Endian: TEndian = endDefault): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadCurrency_LE(var Src: Pointer; out Value: Currency; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadCurrency_LE(Src: Pointer; out Value: Currency): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadCurrency_BE(var Src: Pointer; out Value: Currency; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadCurrency_BE(Src: Pointer; out Value: Currency): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadCurrency(var Src: Pointer; out Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadCurrency(Src: Pointer; out Value: Currency; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetCurrency_LE(var Src: Pointer; Advance: Boolean): Currency; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetCurrency_LE(Src: Pointer): Currency; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetCurrency_BE(var Src: Pointer; Advance: Boolean): Currency; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetCurrency_BE(Src: Pointer): Currency; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetCurrency(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Currency; overload;
Function Ptr_GetCurrency(Src: Pointer; Endian: TEndian = endDefault): Currency; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

Function Ptr_ReadAnsiChar_LE(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_ReadAnsiChar_LE(Src: Pointer; out Value: AnsiChar): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_ReadAnsiChar_BE(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadAnsiChar_BE(Src: Pointer; out Value: AnsiChar): TMemSize; overload;

Function Ptr_ReadAnsiChar(var Src: Pointer; out Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadAnsiChar(Src: Pointer; out Value: AnsiChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetAnsiChar_LE(var Src: Pointer; Advance: Boolean): AnsiChar; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetAnsiChar_LE(Src: Pointer): AnsiChar; overload;
{<lite-end-ln>}

Function Ptr_GetAnsiChar_BE(var Src: Pointer; Advance: Boolean): AnsiChar; overload;
Function Ptr_GetAnsiChar_BE(Src: Pointer): AnsiChar; overload;

Function Ptr_GetAnsiChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): AnsiChar; overload;
Function Ptr_GetAnsiChar(Src: Pointer; Endian: TEndian = endDefault): AnsiChar; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUTF8Char_LE(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_ReadUTF8Char_LE(Src: Pointer; out Value: UTF8Char): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_ReadUTF8Char_BE(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUTF8Char_BE(Src: Pointer; out Value: UTF8Char): TMemSize; overload;

Function Ptr_ReadUTF8Char(var Src: Pointer; out Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUTF8Char(Src: Pointer; out Value: UTF8Char; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetUTF8Char_LE(var Src: Pointer; Advance: Boolean): UTF8Char; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetUTF8Char_LE(Src: Pointer): UTF8Char; overload;
{<lite-end-ln>}

Function Ptr_GetUTF8Char_BE(var Src: Pointer; Advance: Boolean): UTF8Char; overload;
Function Ptr_GetUTF8Char_BE(Src: Pointer): UTF8Char; overload;

Function Ptr_GetUTF8Char(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UTF8Char; overload;
Function Ptr_GetUTF8Char(Src: Pointer; Endian: TEndian = endDefault): UTF8Char; overload;
 
{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadWideChar_LE(var Src: Pointer; out Value: WideChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadWideChar_LE(Src: Pointer; out Value: WideChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadWideChar_BE(var Src: Pointer; out Value: WideChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadWideChar_BE(Src: Pointer; out Value: WideChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadWideChar(var Src: Pointer; out Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadWideChar(Src: Pointer; out Value: WideChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetWideChar_LE(var Src: Pointer; Advance: Boolean): WideChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetWideChar_LE(Src: Pointer): WideChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetWideChar_BE(var Src: Pointer; Advance: Boolean): WideChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetWideChar_BE(Src: Pointer): WideChar; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetWideChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): WideChar; overload;
Function Ptr_GetWideChar(Src: Pointer; Endian: TEndian = endDefault): WideChar; overload;
 
{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeChar_LE(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUnicodeChar_LE(Src: Pointer; out Value: UnicodeChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadUnicodeChar_BE(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUnicodeChar_BE(Src: Pointer; out Value: UnicodeChar): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUnicodeChar(var Src: Pointer; out Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUnicodeChar(Src: Pointer; out Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetUnicodeChar_LE(var Src: Pointer; Advance: Boolean): UnicodeChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUnicodeChar_LE(Src: Pointer): UnicodeChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetUnicodeChar_BE(var Src: Pointer; Advance: Boolean): UnicodeChar; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUnicodeChar_BE(Src: Pointer): UnicodeChar; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUnicodeChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UnicodeChar; overload;
Function Ptr_GetUnicodeChar(Src: Pointer; Endian: TEndian = endDefault): UnicodeChar; overload;
   
{<lite-begin>}
{<lite-block-replace-begin _LE "">}   
//------------------------------------------------------------------------------

Function Ptr_ReadUCS4Char_LE(var Src: Pointer; out Value: UCS4Char; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUCS4Char_LE(Src: Pointer; out Value: UCS4Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadUCS4Char_BE(var Src: Pointer; out Value: UCS4Char; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUCS4Char_BE(Src: Pointer; out Value: UCS4Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUCS4Char(var Src: Pointer; out Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUCS4Char(Src: Pointer; out Value: UCS4Char; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetUCS4Char_LE(var Src: Pointer; Advance: Boolean): UCS4Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUCS4Char_LE(Src: Pointer): UCS4Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetUCS4Char_BE(var Src: Pointer; Advance: Boolean): UCS4Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUCS4Char_BE(Src: Pointer): UCS4Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUCS4Char(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UCS4Char; overload;
Function Ptr_GetUCS4Char(Src: Pointer; Endian: TEndian = endDefault): UCS4Char; overload;
    
{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadChar_LE(var Src: Pointer; out Value: Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadChar_LE(Src: Pointer; out Value: Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadChar_BE(var Src: Pointer; out Value: Char; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadChar_BE(Src: Pointer; out Value: Char): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadChar(var Src: Pointer; out Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadChar(Src: Pointer; out Value: Char; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetChar_LE(var Src: Pointer; Advance: Boolean): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetChar_LE(Src: Pointer): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetChar_BE(var Src: Pointer; Advance: Boolean): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetChar_BE(Src: Pointer): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetChar(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetChar(Src: Pointer; Endian: TEndian = endDefault): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

Function Ptr_ReadShortString_LE(var Src: Pointer; out Value: ShortString; Advance: Boolean): TMemSize; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_ReadShortString_LE(Src: Pointer; out Value: ShortString): TMemSize; overload;
{<lite-end-ln>}

Function Ptr_ReadShortString_BE(var Src: Pointer; out Value: ShortString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadShortString_BE(Src: Pointer; out Value: ShortString): TMemSize; overload;

Function Ptr_ReadShortString(var Src: Pointer; out Value: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadShortString(Src: Pointer; out Value: ShortString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetShortString_LE(var Src: Pointer; Advance: Boolean): ShortString; overload;
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Ptr_GetShortString_LE(Src: Pointer): ShortString; overload;
{<lite-end-ln>}

Function Ptr_GetShortString_BE(var Src: Pointer; Advance: Boolean): ShortString; overload;
Function Ptr_GetShortString_BE(Src: Pointer): ShortString; overload;

Function Ptr_GetShortString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): ShortString; overload;
Function Ptr_GetShortString(Src: Pointer; Endian: TEndian = endDefault): ShortString; overload;
       
{<lite-begin>}
{<lite-block-replace-begin _LE "">}    
//------------------------------------------------------------------------------

Function Ptr_ReadAnsiString_LE(var Src: Pointer; out Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadAnsiString_LE(Src: Pointer; out Value: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadAnsiString_BE(var Src: Pointer; out Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadAnsiString_BE(Src: Pointer; out Value: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadAnsiString(var Src: Pointer; out Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadAnsiString(Src: Pointer; out Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetAnsiString_LE(var Src: Pointer; Advance: Boolean): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetAnsiString_LE(Src: Pointer): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetAnsiString_BE(var Src: Pointer; Advance: Boolean): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetAnsiString_BE(Src: Pointer): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetAnsiString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): AnsiString; overload;
Function Ptr_GetAnsiString(Src: Pointer; Endian: TEndian = endDefault): AnsiString; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUTF8String_LE(var Src: Pointer; out Value: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUTF8String_LE(Src: Pointer; out Value: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadUTF8String_BE(var Src: Pointer; out Value: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUTF8String_BE(Src: Pointer; out Value: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUTF8String(var Src: Pointer; out Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUTF8String(Src: Pointer; out Value: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetUTF8String_LE(var Src: Pointer; Advance: Boolean): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUTF8String_LE(Src: Pointer): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetUTF8String_BE(var Src: Pointer; Advance: Boolean): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUTF8String_BE(Src: Pointer): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUTF8String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UTF8String; overload;
Function Ptr_GetUTF8String(Src: Pointer; Endian: TEndian = endDefault): UTF8String; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadWideString_LE(var Src: Pointer; out Value: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadWideString_LE(Src: Pointer; out Value: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadWideString_BE(var Src: Pointer; out Value: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadWideString_BE(Src: Pointer; out Value: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadWideString(var Src: Pointer; out Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadWideString(Src: Pointer; out Value: WideString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetWideString_LE(var Src: Pointer; Advance: Boolean): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetWideString_LE(Src: Pointer): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetWideString_BE(var Src: Pointer; Advance: Boolean): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetWideString_BE(Src: Pointer): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetWideString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): WideString; overload;
Function Ptr_GetWideString(Src: Pointer; Endian: TEndian = endDefault): WideString; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeString_LE(var Src: Pointer; out Value: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUnicodeString_LE(Src: Pointer; out Value: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadUnicodeString_BE(var Src: Pointer; out Value: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUnicodeString_BE(Src: Pointer; out Value: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUnicodeString(var Src: Pointer; out Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUnicodeString(Src: Pointer; out Value: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetUnicodeString_LE(var Src: Pointer; Advance: Boolean): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUnicodeString_LE(Src: Pointer): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetUnicodeString_BE(var Src: Pointer; Advance: Boolean): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUnicodeString_BE(Src: Pointer): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUnicodeString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString; overload;
Function Ptr_GetUnicodeString(Src: Pointer; Endian: TEndian = endDefault): UnicodeString; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//------------------------------------------------------------------------------

Function Ptr_ReadUCS4String_LE(var Src: Pointer; out Value: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUCS4String_LE(Src: Pointer; out Value: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadUCS4String_BE(var Src: Pointer; out Value: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadUCS4String_BE(Src: Pointer; out Value: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadUCS4String(var Src: Pointer; out Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadUCS4String(Src: Pointer; out Value: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetUCS4String_LE(var Src: Pointer; Advance: Boolean): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUCS4String_LE(Src: Pointer): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetUCS4String_BE(var Src: Pointer; Advance: Boolean): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetUCS4String_BE(Src: Pointer): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetUCS4String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UCS4String; overload;
Function Ptr_GetUCS4String(Src: Pointer; Endian: TEndian = endDefault): UCS4String; overload;
      
{<lite-begin>}
{<lite-block-replace-begin _LE "">}      
//------------------------------------------------------------------------------

Function Ptr_ReadString_LE(var Src: Pointer; out Value: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadString_LE(Src: Pointer; out Value: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_ReadString_BE(var Src: Pointer; out Value: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadString_BE(Src: Pointer; out Value: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadString(var Src: Pointer; out Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadString(Src: Pointer; out Value: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
Function Ptr_GetString_LE(var Src: Pointer; Advance: Boolean): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetString_LE(Src: Pointer): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Ptr_GetString_BE(var Src: Pointer; Advance: Boolean): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetString_BE(Src: Pointer): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetString(Src: Pointer; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Small strings
-------------------------------------------------------------------------------}

Function Ptr_ReadSmallAnsiString_LE(var Src: Pointer; out Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallAnsiString_LE(Src: Pointer; out Value: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallAnsiString_BE(var Src: Pointer; out Value: AnsiString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallAnsiString_BE(Src: Pointer; out Value: AnsiString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallAnsiString(var Src: Pointer; out Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadSmallAnsiString(Src: Pointer; out Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetSmallAnsiString_LE(var Src: Pointer; Advance: Boolean): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallAnsiString_LE(Src: Pointer): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallAnsiString_BE(var Src: Pointer; Advance: Boolean): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallAnsiString_BE(Src: Pointer): AnsiString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallAnsiString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): AnsiString; overload;
Function Ptr_GetSmallAnsiString(Src: Pointer; Endian: TEndian = endDefault): AnsiString; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUTF8String_LE(var Src: Pointer; out Value: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallUTF8String_LE(Src: Pointer; out Value: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallUTF8String_BE(var Src: Pointer; out Value: UTF8String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallUTF8String_BE(Src: Pointer; out Value: UTF8String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallUTF8String(var Src: Pointer; out Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadSmallUTF8String(Src: Pointer; out Value: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetSmallUTF8String_LE(var Src: Pointer; Advance: Boolean): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallUTF8String_LE(Src: Pointer): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallUTF8String_BE(var Src: Pointer; Advance: Boolean): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallUTF8String_BE(Src: Pointer): UTF8String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallUTF8String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UTF8String; overload;
Function Ptr_GetSmallUTF8String(Src: Pointer; Endian: TEndian = endDefault): UTF8String; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallWideString_LE(var Src: Pointer; out Value: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallWideString_LE(Src: Pointer; out Value: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallWideString_BE(var Src: Pointer; out Value: WideString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallWideString_BE(Src: Pointer; out Value: WideString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallWideString(var Src: Pointer; out Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadSmallWideString(Src: Pointer; out Value: WideString; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetSmallWideString_LE(var Src: Pointer; Advance: Boolean): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallWideString_LE(Src: Pointer): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallWideString_BE(var Src: Pointer; Advance: Boolean): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallWideString_BE(Src: Pointer): WideString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallWideString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): WideString; overload;
Function Ptr_GetSmallWideString(Src: Pointer; Endian: TEndian = endDefault): WideString; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUnicodeString_LE(var Src: Pointer; out Value: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallUnicodeString_LE(Src: Pointer; out Value: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallUnicodeString_BE(var Src: Pointer; out Value: UnicodeString; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallUnicodeString_BE(Src: Pointer; out Value: UnicodeString): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallUnicodeString(var Src: Pointer; out Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadSmallUnicodeString(Src: Pointer; out Value: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetSmallUnicodeString_LE(var Src: Pointer; Advance: Boolean): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallUnicodeString_LE(Src: Pointer): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallUnicodeString_BE(var Src: Pointer; Advance: Boolean): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallUnicodeString_BE(Src: Pointer): UnicodeString; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallUnicodeString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString; overload;
Function Ptr_GetSmallUnicodeString(Src: Pointer; Endian: TEndian = endDefault): UnicodeString; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUCS4String_LE(var Src: Pointer; out Value: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallUCS4String_LE(Src: Pointer; out Value: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallUCS4String_BE(var Src: Pointer; out Value: UCS4String; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadSmallUCS4String_BE(Src: Pointer; out Value: UCS4String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallUCS4String(var Src: Pointer; out Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Ptr_ReadSmallUCS4String(Src: Pointer; out Value: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

Function Ptr_GetSmallUCS4String_LE(var Src: Pointer; Advance: Boolean): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallUCS4String_LE(Src: Pointer): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallUCS4String_BE(var Src: Pointer; Advance: Boolean): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallUCS4String_BE(Src: Pointer): UCS4String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallUCS4String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UCS4String; overload;
Function Ptr_GetSmallUCS4String(Src: Pointer; Endian: TEndian = endDefault): UCS4String; overload;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallString_LE(var Src: Pointer; out Value: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadSmallString_LE(Src: Pointer; out Value: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallString_BE(var Src: Pointer; out Value: String; Advance: Boolean): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadSmallString_BE(Src: Pointer; out Value: String): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_ReadSmallString(var Src: Pointer; out Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_ReadSmallString(Src: Pointer; out Value: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallString_LE(var Src: Pointer; Advance: Boolean): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallString_LE(Src: Pointer): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallString_BE(var Src: Pointer; Advance: Boolean): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallString_BE(Src: Pointer): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Ptr_GetSmallString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Ptr_GetSmallString(Src: Pointer; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

Function Ptr_ReadBuffer_LE(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize; overload;
Function Ptr_ReadBuffer_LE(Src: Pointer; out Buffer; Size: TMemSize): TMemSize; overload;
{<lite-end-ln>}

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

{<lite-begin>}
{===============================================================================
--------------------------------------------------------------------------------
                                 Stream writing
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_WriteBool_LE(Stream: TStream; Value: ByteBool; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteBool_BE(Stream: TStream; Value: ByteBool; Advance: Boolean = True): TMemSize;

Function Stream_WriteBool(Stream: TStream; Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteBool(Stream: TStream; Value: ByteBool; Endian: TEndian = endDefault): TMemSize; overload;
            
{<lite-begin>}                
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteBoolean_LE(Stream: TStream; Value: Boolean; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_WriteBoolean_BE(Stream: TStream; Value: Boolean; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteBoolean(Stream: TStream; Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteBoolean(Stream: TStream; Value: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>} 
{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_WriteInt8_LE(Stream: TStream; Value: Int8; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteInt8_BE(Stream: TStream; Value: Int8; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt8(Stream: TStream; Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteInt8(Stream: TStream; Value: Int8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUInt8_LE(Stream: TStream; Value: UInt8; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUInt8_BE(Stream: TStream; Value: UInt8; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt8(Stream: TStream; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUInt8(Stream: TStream; Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteInt16_LE(Stream: TStream; Value: Int16; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteInt16_BE(Stream: TStream; Value: Int16; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt16(Stream: TStream; Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteInt16(Stream: TStream; Value: Int16; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------
 
{<lite-replace _LE "">}
Function Stream_WriteUInt16_LE(Stream: TStream; Value: UInt16; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUInt16_BE(Stream: TStream; Value: UInt16; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt16(Stream: TStream; Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUInt16(Stream: TStream; Value: UInt16; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteInt32_LE(Stream: TStream; Value: Int32; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteInt32_BE(Stream: TStream; Value: Int32; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt32(Stream: TStream; Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteInt32(Stream: TStream; Value: Int32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUInt32_LE(Stream: TStream; Value: UInt32; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUInt32_BE(Stream: TStream; Value: UInt32; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt32(Stream: TStream; Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUInt32(Stream: TStream; Value: UInt32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteInt64_LE(Stream: TStream; Value: Int64; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteInt64_BE(Stream: TStream; Value: Int64; Advance: Boolean = True): TMemSize;

Function Stream_WriteInt64(Stream: TStream; Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteInt64(Stream: TStream; Value: Int64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUInt64_LE(Stream: TStream; Value: UInt64; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUInt64_BE(Stream: TStream; Value: UInt64; Advance: Boolean = True): TMemSize;

Function Stream_WriteUInt64(Stream: TStream; Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUInt64(Stream: TStream; Value: UInt64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_WriteFloat32_LE(Stream: TStream; Value: Float32; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteFloat32_BE(Stream: TStream; Value: Float32; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat32(Stream: TStream; Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteFloat32(Stream: TStream; Value: Float32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteFloat64_LE(Stream: TStream; Value: Float64; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteFloat64_BE(Stream: TStream; Value: Float64; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat64(Stream: TStream; Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteFloat64(Stream: TStream; Value: Float64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------
 
{<lite-replace _LE "">}
Function Stream_WriteFloat80_LE(Stream: TStream; Value: Float80; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteFloat80_BE(Stream: TStream; Value: Float80; Advance: Boolean = True): TMemSize;

Function Stream_WriteFloat80(Stream: TStream; Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteFloat80(Stream: TStream; Value: Float80; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteDateTime_LE(Stream: TStream; Value: TDateTime; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_WriteDateTime_BE(Stream: TStream; Value: TDateTime; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteDateTime(Stream: TStream; Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteDateTime(Stream: TStream; Value: TDateTime; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteCurrency_LE(Stream: TStream; Value: Currency; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteCurrency_BE(Stream: TStream; Value: Currency; Advance: Boolean = True): TMemSize;

Function Stream_WriteCurrency(Stream: TStream; Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteCurrency(Stream: TStream; Value: Currency; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_WriteAnsiChar_LE(Stream: TStream; Value: AnsiChar; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteAnsiChar_BE(Stream: TStream; Value: AnsiChar; Advance: Boolean = True): TMemSize;

Function Stream_WriteAnsiChar(Stream: TStream; Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteAnsiChar(Stream: TStream; Value: AnsiChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUTF8Char_LE(Stream: TStream; Value: UTF8Char; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUTF8Char_BE(Stream: TStream; Value: UTF8Char; Advance: Boolean = True): TMemSize;

Function Stream_WriteUTF8Char(Stream: TStream; Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUTF8Char(Stream: TStream; Value: UTF8Char; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteWideChar_LE(Stream: TStream; Value: WideChar; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteWideChar_BE(Stream: TStream; Value: WideChar; Advance: Boolean = True): TMemSize;

Function Stream_WriteWideChar(Stream: TStream; Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteWideChar(Stream: TStream; Value: WideChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUnicodeChar_LE(Stream: TStream; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUnicodeChar_BE(Stream: TStream; Value: UnicodeChar; Advance: Boolean = True): TMemSize;

Function Stream_WriteUnicodeChar(Stream: TStream; Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUnicodeChar(Stream: TStream; Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUCS4Char_LE(Stream: TStream; Value: UCS4Char; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUCS4Char_BE(Stream: TStream; Value: UCS4Char; Advance: Boolean = True): TMemSize;

Function Stream_WriteUCS4Char(Stream: TStream; Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUCS4Char(Stream: TStream; Value: UCS4Char; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteChar_LE(Stream: TStream; Value: Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_WriteChar_BE(Stream: TStream; Value: Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteChar(Stream: TStream; Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteChar(Stream: TStream; Value: Char; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_WriteShortString_LE(Stream: TStream; const Value: ShortString; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteShortString_BE(Stream: TStream; const Value: ShortString; Advance: Boolean = True): TMemSize;

Function Stream_WriteShortString(Stream: TStream; const Value: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteShortString(Stream: TStream; const Value: ShortString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteAnsiString_LE(Stream: TStream; const Value: AnsiString; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteAnsiString_BE(Stream: TStream; const Value: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_WriteAnsiString(Stream: TStream; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteAnsiString(Stream: TStream; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUTF8String_LE(Stream: TStream; const Value: UTF8String; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUTF8String_BE(Stream: TStream; const Value: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_WriteUTF8String(Stream: TStream; const Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUTF8String(Stream: TStream; const Value: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteWideString_LE(Stream: TStream; const Value: WideString; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteWideString_BE(Stream: TStream; const Value: WideString; Advance: Boolean = True): TMemSize;

Function Stream_WriteWideString(Stream: TStream; const Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteWideString(Stream: TStream; const Value: WideString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUnicodeString_LE(Stream: TStream; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUnicodeString_BE(Stream: TStream; const Value: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_WriteUnicodeString(Stream: TStream; const Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUnicodeString(Stream: TStream; const Value: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteUCS4String_LE(Stream: TStream; const Value: UCS4String; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteUCS4String_BE(Stream: TStream; const Value: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_WriteUCS4String(Stream: TStream; const Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteUCS4String(Stream: TStream; const Value: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteString_LE(Stream: TStream; const Value: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_WriteString_BE(Stream: TStream; const Value: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteString(Stream: TStream; const Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteString(Stream: TStream; const Value: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Small strings
-------------------------------------------------------------------------------}

Function Stream_WriteSmallAnsiString_LE(Stream: TStream; const Value: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallAnsiString_BE(Stream: TStream; const Value: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallAnsiString(Stream: TStream; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteSmallAnsiString(Stream: TStream; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUTF8String_LE(Stream: TStream; const Value: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallUTF8String_BE(Stream: TStream; const Value: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallUTF8String(Stream: TStream; const Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteSmallUTF8String(Stream: TStream; const Value: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteSmallWideString_LE(Stream: TStream; const Value: WideString; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallWideString_BE(Stream: TStream; const Value: WideString; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallWideString(Stream: TStream; const Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteSmallWideString(Stream: TStream; const Value: WideString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUnicodeString_LE(Stream: TStream; const Value: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallUnicodeString_BE(Stream: TStream; const Value: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallUnicodeString(Stream: TStream; const Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteSmallUnicodeString(Stream: TStream; const Value: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUCS4String_LE(Stream: TStream; const Value: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallUCS4String_BE(Stream: TStream; const Value: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_WriteSmallUCS4String(Stream: TStream; const Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteSmallUCS4String(Stream: TStream; const Value: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

//------------------------------------------------------------------------------

Function Stream_WriteSmallString_LE(Stream: TStream; const Value: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteSmallString_BE(Stream: TStream; const Value: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_WriteSmallString(Stream: TStream; const Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_WriteSmallString(Stream: TStream; const Value: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_WriteBuffer_LE(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteBuffer_BE(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;

Function Stream_WriteBuffer(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteBuffer(Stream: TStream; const Buffer; Size: TMemSize; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_WriteBytes_LE(Stream: TStream; const Value: array of UInt8; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_WriteBytes_BE(Stream: TStream; const Value: array of UInt8; Advance: Boolean = True): TMemSize;

Function Stream_WriteBytes(Stream: TStream; const Value: array of UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteBytes(Stream: TStream; const Value: array of UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_FillBytes_LE(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_FillBytes_BE(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;

Function Stream_FillBytes(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_FillBytes(Stream: TStream; Count: TMemSize; Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{-------------------------------------------------------------------------------
    Variants
-------------------------------------------------------------------------------}

Function Stream_WriteVariant_LE(Stream: TStream; const Value: Variant; Advance: Boolean = True): TMemSize;

Function Stream_WriteVariant_BE(Stream: TStream; const Value: Variant; Advance: Boolean = True): TMemSize;

Function Stream_WriteVariant(Stream: TStream; const Value: Variant; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_WriteVariant(Stream: TStream; const Value: Variant; Endian: TEndian = endDefault): TMemSize; overload;


{<lite-begin>}
{===============================================================================
--------------------------------------------------------------------------------
                                 Stream reading
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_ReadBool_LE(Stream: TStream; out Value: ByteBool; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadBool_BE(Stream: TStream; out Value: ByteBool; Advance: Boolean = True): TMemSize;

Function Stream_ReadBool(Stream: TStream; out Value: ByteBool; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadBool(Stream: TStream; out Value: ByteBool; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Stream_GetBool_LE(Stream: TStream; Advance: Boolean = True): ByteBool;
{<lite-end-ln>}

Function Stream_GetBool_BE(Stream: TStream; Advance: Boolean = True): ByteBool;

Function Stream_GetBool(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): ByteBool; overload;
Function Stream_GetBool(Stream: TStream; Endian: TEndian = endDefault): ByteBool; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadBoolean_LE(Stream: TStream; out Value: Boolean; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_ReadBoolean_BE(Stream: TStream; out Value: Boolean; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadBoolean(Stream: TStream; out Value: Boolean; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadBoolean(Stream: TStream; out Value: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetBoolean_LE(Stream: TStream; Advance: Boolean = True): Boolean;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetBoolean_BE(Stream: TStream; Advance: Boolean = True): Boolean;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetBoolean(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetBoolean(Stream: TStream; Endian: TEndian = endDefault): Boolean; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_ReadInt8_LE(Stream: TStream; out Value: Int8; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadInt8_BE(Stream: TStream; out Value: Int8; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt8(Stream: TStream; out Value: Int8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadInt8(Stream: TStream; out Value: Int8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Stream_GetInt8_LE(Stream: TStream; Advance: Boolean = True): Int8;
{<lite-end-ln>}

Function Stream_GetInt8_BE(Stream: TStream; Advance: Boolean = True): Int8;

Function Stream_GetInt8(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int8; overload;
Function Stream_GetInt8(Stream: TStream; Endian: TEndian = endDefault): Int8; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUInt8_LE(Stream: TStream; out Value: UInt8; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUInt8_BE(Stream: TStream; out Value: UInt8; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt8(Stream: TStream; out Value: UInt8; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUInt8(Stream: TStream; out Value: UInt8; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Stream_GetUInt8_LE(Stream: TStream; Advance: Boolean = True): UInt8;
{<lite-end-ln>}

Function Stream_GetUInt8_BE(Stream: TStream; Advance: Boolean = True): UInt8;

Function Stream_GetUInt8(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt8; overload;
Function Stream_GetUInt8(Stream: TStream; Endian: TEndian = endDefault): UInt8; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadInt16_LE(Stream: TStream; out Value: Int16; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadInt16_BE(Stream: TStream; out Value: Int16; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt16(Stream: TStream; out Value: Int16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadInt16(Stream: TStream; out Value: Int16; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetInt16_LE(Stream: TStream; Advance: Boolean = True): Int16;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetInt16_BE(Stream: TStream; Advance: Boolean = True): Int16;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt16(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int16; overload;
Function Stream_GetInt16(Stream: TStream; Endian: TEndian = endDefault): Int16; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUInt16_LE(Stream: TStream; out Value: UInt16; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUInt16_BE(Stream: TStream; out Value: UInt16; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt16(Stream: TStream; out Value: UInt16; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUInt16(Stream: TStream; out Value: UInt16; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetUInt16_LE(Stream: TStream; Advance: Boolean = True): UInt16;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetUInt16_BE(Stream: TStream; Advance: Boolean = True): UInt16;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt16(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt16; overload;
Function Stream_GetUInt16(Stream: TStream; Endian: TEndian = endDefault): UInt16; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadInt32_LE(Stream: TStream; out Value: Int32; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadInt32_BE(Stream: TStream; out Value: Int32; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt32(Stream: TStream; out Value: Int32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadInt32(Stream: TStream; out Value: Int32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetInt32_LE(Stream: TStream; Advance: Boolean = True): Int32;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetInt32_BE(Stream: TStream; Advance: Boolean = True): Int32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int32; overload;
Function Stream_GetInt32(Stream: TStream; Endian: TEndian = endDefault): Int32; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUInt32_LE(Stream: TStream; out Value: UInt32; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUInt32_BE(Stream: TStream; out Value: UInt32; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt32(Stream: TStream; out Value: UInt32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUInt32(Stream: TStream; out Value: UInt32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetUInt32_LE(Stream: TStream; Advance: Boolean = True): UInt32;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetUInt32_BE(Stream: TStream; Advance: Boolean = True): UInt32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt32; overload;
Function Stream_GetUInt32(Stream: TStream; Endian: TEndian = endDefault): UInt32; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadInt64_LE(Stream: TStream; out Value: Int64; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadInt64_BE(Stream: TStream; out Value: Int64; Advance: Boolean = True): TMemSize;

Function Stream_ReadInt64(Stream: TStream; out Value: Int64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadInt64(Stream: TStream; out Value: Int64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetInt64_LE(Stream: TStream; Advance: Boolean = True): Int64;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetInt64_BE(Stream: TStream; Advance: Boolean = True): Int64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetInt64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Int64; overload;
Function Stream_GetInt64(Stream: TStream; Endian: TEndian = endDefault): Int64; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUInt64_LE(Stream: TStream; out Value: UInt64; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUInt64_BE(Stream: TStream; out Value: UInt64; Advance: Boolean = True): TMemSize;

Function Stream_ReadUInt64(Stream: TStream; out Value: UInt64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUInt64(Stream: TStream; out Value: UInt64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetUInt64_LE(Stream: TStream; Advance: Boolean = True): UInt64;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetUInt64_BE(Stream: TStream; Advance: Boolean = True): UInt64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUInt64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UInt64; overload;
Function Stream_GetUInt64(Stream: TStream; Endian: TEndian = endDefault): UInt64; overload;

{<lite-begin>}
{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_ReadFloat32_LE(Stream: TStream; out Value: Float32; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadFloat32_BE(Stream: TStream; out Value: Float32; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat32(Stream: TStream; out Value: Float32; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadFloat32(Stream: TStream; out Value: Float32; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetFloat32_LE(Stream: TStream; Advance: Boolean = True): Float32;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetFloat32_BE(Stream: TStream; Advance: Boolean = True): Float32;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat32(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float32; overload;
Function Stream_GetFloat32(Stream: TStream; Endian: TEndian = endDefault): Float32; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadFloat64_LE(Stream: TStream; out Value: Float64; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadFloat64_BE(Stream: TStream; out Value: Float64; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat64(Stream: TStream; out Value: Float64; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadFloat64(Stream: TStream; out Value: Float64; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetFloat64_LE(Stream: TStream; Advance: Boolean = True): Float64;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetFloat64_BE(Stream: TStream; Advance: Boolean = True): Float64;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat64(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float64; overload;
Function Stream_GetFloat64(Stream: TStream; Endian: TEndian = endDefault): Float64; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadFloat80_LE(Stream: TStream; out Value: Float80; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadFloat80_BE(Stream: TStream; out Value: Float80; Advance: Boolean = True): TMemSize;

Function Stream_ReadFloat80(Stream: TStream; out Value: Float80; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadFloat80(Stream: TStream; out Value: Float80; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetFloat80_LE(Stream: TStream; Advance: Boolean = True): Float80;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetFloat80_BE(Stream: TStream; Advance: Boolean = True): Float80;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetFloat80(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Float80; overload;
Function Stream_GetFloat80(Stream: TStream; Endian: TEndian = endDefault): Float80; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadDateTime_LE(Stream: TStream; out Value: TDateTime; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_ReadDateTime_BE(Stream: TStream; out Value: TDateTime; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadDateTime(Stream: TStream; out Value: TDateTime; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadDateTime(Stream: TStream; out Value: TDateTime; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetDateTime_LE(Stream: TStream; Advance: Boolean = True): TDateTime;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetDateTime_BE(Stream: TStream; Advance: Boolean = True): TDateTime;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetDateTime(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetDateTime(Stream: TStream; Endian: TEndian = endDefault): TDateTime; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadCurrency_LE(Stream: TStream; out Value: Currency; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadCurrency_BE(Stream: TStream; out Value: Currency; Advance: Boolean = True): TMemSize;

Function Stream_ReadCurrency(Stream: TStream; out Value: Currency; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadCurrency(Stream: TStream; out Value: Currency; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetCurrency_LE(Stream: TStream; Advance: Boolean = True): Currency;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetCurrency_BE(Stream: TStream; Advance: Boolean = True): Currency;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetCurrency(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Currency; overload;
Function Stream_GetCurrency(Stream: TStream; Endian: TEndian = endDefault): Currency; overload;

{<lite-begin>}
{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_ReadAnsiChar_LE(Stream: TStream; out Value: AnsiChar; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadAnsiChar_BE(Stream: TStream; out Value: AnsiChar; Advance: Boolean = True): TMemSize;

Function Stream_ReadAnsiChar(Stream: TStream; out Value: AnsiChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadAnsiChar(Stream: TStream; out Value: AnsiChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Stream_GetAnsiChar_LE(Stream: TStream; Advance: Boolean = True): AnsiChar;
{<lite-end-ln>}

Function Stream_GetAnsiChar_BE(Stream: TStream; Advance: Boolean = True): AnsiChar;

Function Stream_GetAnsiChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): AnsiChar; overload;
Function Stream_GetAnsiChar(Stream: TStream; Endian: TEndian = endDefault): AnsiChar; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUTF8Char_LE(Stream: TStream; out Value: UTF8Char; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUTF8Char_BE(Stream: TStream; out Value: UTF8Char; Advance: Boolean = True): TMemSize;

Function Stream_ReadUTF8Char(Stream: TStream; out Value: UTF8Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUTF8Char(Stream: TStream; out Value: UTF8Char; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Stream_GetUTF8Char_LE(Stream: TStream; Advance: Boolean = True): UTF8Char;
{<lite-end-ln>}

Function Stream_GetUTF8Char_BE(Stream: TStream; Advance: Boolean = True): UTF8Char;

Function Stream_GetUTF8Char(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UTF8Char; overload;
Function Stream_GetUTF8Char(Stream: TStream; Endian: TEndian = endDefault): UTF8Char; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadWideChar_LE(Stream: TStream; out Value: WideChar; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadWideChar_BE(Stream: TStream; out Value: WideChar; Advance: Boolean = True): TMemSize;

Function Stream_ReadWideChar(Stream: TStream; out Value: WideChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadWideChar(Stream: TStream; out Value: WideChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetWideChar_LE(Stream: TStream; Advance: Boolean = True): WideChar;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetWideChar_BE(Stream: TStream; Advance: Boolean = True): WideChar;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetWideChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): WideChar; overload;
Function Stream_GetWideChar(Stream: TStream; Endian: TEndian = endDefault): WideChar; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUnicodeChar_LE(Stream: TStream; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUnicodeChar_BE(Stream: TStream; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;

Function Stream_ReadUnicodeChar(Stream: TStream; out Value: UnicodeChar; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUnicodeChar(Stream: TStream; out Value: UnicodeChar; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetUnicodeChar_LE(Stream: TStream; Advance: Boolean = True): UnicodeChar;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetUnicodeChar_BE(Stream: TStream; Advance: Boolean = True): UnicodeChar;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUnicodeChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UnicodeChar; overload;
Function Stream_GetUnicodeChar(Stream: TStream; Endian: TEndian = endDefault): UnicodeChar; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUCS4Char_LE(Stream: TStream; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUCS4Char_BE(Stream: TStream; out Value: UCS4Char; Advance: Boolean = True): TMemSize;

Function Stream_ReadUCS4Char(Stream: TStream; out Value: UCS4Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUCS4Char(Stream: TStream; out Value: UCS4Char; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetUCS4Char_LE(Stream: TStream; Advance: Boolean = True): UCS4Char;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetUCS4Char_BE(Stream: TStream; Advance: Boolean = True): UCS4Char;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUCS4Char(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UCS4Char; overload;
Function Stream_GetUCS4Char(Stream: TStream; Endian: TEndian = endDefault): UCS4Char; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadChar_LE(Stream: TStream; out Value: Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_ReadChar_BE(Stream: TStream; out Value: Char; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadChar(Stream: TStream; out Value: Char; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadChar(Stream: TStream; out Value: Char; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetChar_LE(Stream: TStream; Advance: Boolean = True): Char;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetChar_BE(Stream: TStream; Advance: Boolean = True): Char;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetChar(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetChar(Stream: TStream; Endian: TEndian = endDefault): Char; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_ReadShortString_LE(Stream: TStream; out Value: ShortString; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadShortString_BE(Stream: TStream; out Value: ShortString; Advance: Boolean = True): TMemSize;

Function Stream_ReadShortString(Stream: TStream; out Value: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadShortString(Stream: TStream; out Value: ShortString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
{<lite-append "\x7B$IFDEF CanInline\x7D inline;\x7B$ENDIF\x7D">}
Function Stream_GetShortString_LE(Stream: TStream; Advance: Boolean = True): ShortString;
{<lite-end-ln>}

Function Stream_GetShortString_BE(Stream: TStream; Advance: Boolean = True): ShortString;

Function Stream_GetShortString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): ShortString; overload;
Function Stream_GetShortString(Stream: TStream; Endian: TEndian = endDefault): ShortString; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadAnsiString_LE(Stream: TStream; out Value: AnsiString; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadAnsiString_BE(Stream: TStream; out Value: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_ReadAnsiString(Stream: TStream; out Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadAnsiString(Stream: TStream; out Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetAnsiString_LE(Stream: TStream; Advance: Boolean = True): AnsiString;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetAnsiString_BE(Stream: TStream; Advance: Boolean = True): AnsiString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetAnsiString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): AnsiString; overload;
Function Stream_GetAnsiString(Stream: TStream; Endian: TEndian = endDefault): AnsiString; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUTF8String_LE(Stream: TStream; out Value: UTF8String; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUTF8String_BE(Stream: TStream; out Value: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_ReadUTF8String(Stream: TStream; out Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUTF8String(Stream: TStream; out Value: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetUTF8String_LE(Stream: TStream; Advance: Boolean = True): UTF8String;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetUTF8String_BE(Stream: TStream; Advance: Boolean = True): UTF8String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUTF8String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UTF8String; overload;
Function Stream_GetUTF8String(Stream: TStream; Endian: TEndian = endDefault): UTF8String; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadWideString_LE(Stream: TStream; out Value: WideString; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadWideString_BE(Stream: TStream; out Value: WideString; Advance: Boolean = True): TMemSize;

Function Stream_ReadWideString(Stream: TStream; out Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadWideString(Stream: TStream; out Value: WideString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetWideString_LE(Stream: TStream; Advance: Boolean = True): WideString;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetWideString_BE(Stream: TStream; Advance: Boolean = True): WideString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetWideString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): WideString; overload;
Function Stream_GetWideString(Stream: TStream; Endian: TEndian = endDefault): WideString; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUnicodeString_LE(Stream: TStream; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUnicodeString_BE(Stream: TStream; out Value: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_ReadUnicodeString(Stream: TStream; out Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUnicodeString(Stream: TStream; out Value: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetUnicodeString_LE(Stream: TStream; Advance: Boolean = True): UnicodeString;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetUnicodeString_BE(Stream: TStream; Advance: Boolean = True): UnicodeString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUnicodeString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString; overload;
Function Stream_GetUnicodeString(Stream: TStream; Endian: TEndian = endDefault): UnicodeString; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadUCS4String_LE(Stream: TStream; out Value: UCS4String; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

Function Stream_ReadUCS4String_BE(Stream: TStream; out Value: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_ReadUCS4String(Stream: TStream; out Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadUCS4String(Stream: TStream; out Value: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetUCS4String_LE(Stream: TStream; Advance: Boolean = True): UCS4String;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetUCS4String_BE(Stream: TStream; Advance: Boolean = True): UCS4String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetUCS4String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UCS4String; overload;
Function Stream_GetUCS4String(Stream: TStream; Endian: TEndian = endDefault): UCS4String; overload;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_ReadString_LE(Stream: TStream; out Value: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_ReadString_BE(Stream: TStream; out Value: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadString(Stream: TStream; out Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadString(Stream: TStream; out Value: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{<lite-replace _LE "">}
Function Stream_GetString_LE(Stream: TStream; Advance: Boolean = True): String;{$IFDEF CanInline} inline;{$ENDIF}
{<lite-end-ln>}

Function Stream_GetString_BE(Stream: TStream; Advance: Boolean = True): String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetString(Stream: TStream; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

{-------------------------------------------------------------------------------
    Small strings
-------------------------------------------------------------------------------}

Function Stream_ReadSmallAnsiString_LE(Stream: TStream; out Value: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallAnsiString_BE(Stream: TStream; out Value: AnsiString; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallAnsiString(Stream: TStream; out Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadSmallAnsiString(Stream: TStream; out Value: AnsiString; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetSmallAnsiString_LE(Stream: TStream; Advance: Boolean = True): AnsiString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallAnsiString_BE(Stream: TStream; Advance: Boolean = True): AnsiString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallAnsiString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): AnsiString; overload;
Function Stream_GetSmallAnsiString(Stream: TStream; Endian: TEndian = endDefault): AnsiString; overload;

//------------------------------------------------------------------------------

Function Stream_ReadSmallUTF8String_LE(Stream: TStream; out Value: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallUTF8String_BE(Stream: TStream; out Value: UTF8String; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallUTF8String(Stream: TStream; out Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadSmallUTF8String(Stream: TStream; out Value: UTF8String; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetSmallUTF8String_LE(Stream: TStream; Advance: Boolean = True): UTF8String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallUTF8String_BE(Stream: TStream; Advance: Boolean = True): UTF8String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallUTF8String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UTF8String; overload;
Function Stream_GetSmallUTF8String(Stream: TStream; Endian: TEndian = endDefault): UTF8String; overload;

//------------------------------------------------------------------------------

Function Stream_ReadSmallWideString_LE(Stream: TStream; out Value: WideString; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallWideString_BE(Stream: TStream; out Value: WideString; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallWideString(Stream: TStream; out Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadSmallWideString(Stream: TStream; out Value: WideString; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetSmallWideString_LE(Stream: TStream; Advance: Boolean = True): WideString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallWideString_BE(Stream: TStream; Advance: Boolean = True): WideString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallWideString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): WideString; overload;
Function Stream_GetSmallWideString(Stream: TStream; Endian: TEndian = endDefault): WideString; overload;

//------------------------------------------------------------------------------

Function Stream_ReadSmallUnicodeString_LE(Stream: TStream; out Value: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallUnicodeString_BE(Stream: TStream; out Value: UnicodeString; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallUnicodeString(Stream: TStream; out Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadSmallUnicodeString(Stream: TStream; out Value: UnicodeString; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetSmallUnicodeString_LE(Stream: TStream; Advance: Boolean = True): UnicodeString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallUnicodeString_BE(Stream: TStream; Advance: Boolean = True): UnicodeString;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallUnicodeString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString; overload;
Function Stream_GetSmallUnicodeString(Stream: TStream; Endian: TEndian = endDefault): UnicodeString; overload;

//------------------------------------------------------------------------------

Function Stream_ReadSmallUCS4String_LE(Stream: TStream; out Value: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallUCS4String_BE(Stream: TStream; out Value: UCS4String; Advance: Boolean = True): TMemSize;

Function Stream_ReadSmallUCS4String(Stream: TStream; out Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;
Function Stream_ReadSmallUCS4String(Stream: TStream; out Value: UCS4String; Endian: TEndian = endDefault): TMemSize; overload;

Function Stream_GetSmallUCS4String_LE(Stream: TStream; Advance: Boolean = True): UCS4String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallUCS4String_BE(Stream: TStream; Advance: Boolean = True): UCS4String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallUCS4String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UCS4String; overload;
Function Stream_GetSmallUCS4String(Stream: TStream; Endian: TEndian = endDefault): UCS4String; overload;

//------------------------------------------------------------------------------

Function Stream_ReadSmallString_LE(Stream: TStream; out Value: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadSmallString_BE(Stream: TStream; out Value: String; Advance: Boolean = True): TMemSize;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_ReadSmallString(Stream: TStream; out Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_ReadSmallString(Stream: TStream; out Value: String; Endian: TEndian = endDefault): TMemSize; overload;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallString_LE(Stream: TStream; Advance: Boolean = True): String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallString_BE(Stream: TStream; Advance: Boolean = True): String;{$IFDEF CanInline} inline;{$ENDIF}

Function Stream_GetSmallString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}
Function Stream_GetSmallString(Stream: TStream; Endian: TEndian = endDefault): String; overload;{$IFDEF CanInline} inline;{$ENDIF}

{<lite-begin>}
{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
Function Stream_ReadBuffer_LE(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
{<lite-end-ln>}

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


{===============================================================================
--------------------------------------------------------------------------------
                                 TCustomStreamer
--------------------------------------------------------------------------------
===============================================================================}
type
  TBSBookmarkID = type Integer;

  TBSBookmarkData = record
    ID:       TBSBookmarkID;
    Position: Int64;
  end;
  PBSBookmarkData = ^TBSBookmarkData;

{===============================================================================
    TCustomStreamer - class declaration
===============================================================================}
type
  TCustomStreamer = class(TCustomListObject)
  protected
    fEndian:            TEndian;
    fStart:             Int64;
    fBookmarks:         array of TBSBookmarkData;
    fBookmarkCount:     Integer;
    fChanged:           Boolean;
    fChangingCounter:   Integer;
    fOnChangeEvent:     TNotifyEvent;
    fOnChangeCallback:  TNotifyCallback;
    Function GetPosition: Int64; virtual; abstract;
    procedure SetPosition(NewPosition: Int64); virtual; abstract;
    Function GetOffset: Int64; virtual;
    procedure SetOffset(NewOffset: Int64); virtual;
    Function GetBookmark(Index: Integer): TBSBookmarkData; virtual;
    procedure SetBookmark(Index: Integer; Value: TBSBookmarkData); virtual;
    Function GetBookmarkPtr(Index: Integer): PBSBookmarkData; virtual;
    Function GetCapacity: Integer; override;
    procedure SetCapacity(Value: Integer); override;
    Function GetCount: Integer; override;
    procedure SetCount(Value: Integer); override;
    Function WriteOpenByteArray(Ptr: Pointer; Count: TMemSize; Advance: Boolean): TMemSize; virtual;
    Function WriteValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual; abstract;
    Function WriteValueAt(Position: Int64; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    Function WriteValueAtOffset(Offset: Int64; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    Function WriteValueTo(ID: TBSBookmarkID; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    Function WriteValueToIndex(Index: Integer; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    Function ReadValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual; abstract;
    Function ReadValueAt(Position: Int64; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    Function ReadValueAtOffset(Offset: Int64; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    Function ReadValueFrom(ID: TBSBookmarkID; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    Function ReadValueFromIndex(Index: Integer; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    procedure DoChange; virtual;
    procedure Initialize; virtual;
    procedure Finalize; virtual;
  public
    destructor Destroy; override;
    // bookmark list methods
    Function BeginUpdate: Integer; virtual;
    Function EndUpdate: Integer; virtual;
    Function LowIndex: Integer; override;
    Function HighIndex: Integer; override;
    Function BookmarkLowIndex: Integer; virtual;
    Function BookmarkHighIndex: Integer; virtual;
    Function BookmarkCheckIndex(Index: Integer): Boolean; virtual;
    Function BookmarkIndexOf(ID: TBSBookmarkID): Integer; virtual;
    Function BookmarkFind(ID: TBSBookmarkID; out Index: Integer): Boolean; virtual;
    Function BookmarkAdd(ID: TBSBookmarkID; Position: Int64): Integer; overload; virtual;
    Function BookmarkAdd(ID: TBSBookmarkID): Integer; overload; virtual;
    procedure BookmarkInsert(Index: Integer; ID: TBSBookmarkID; Position: Int64); overload; virtual;
    procedure BookmarkInsert(Index: Integer; ID: TBSBookmarkID); overload; virtual;
    procedure BookmarkMove(SrcIndex,DstIndex: Integer); virtual;
    procedure BookmarkExchange(Index1,Index2: Integer); virtual;
    Function BookmarkExtract(ID: TBSBookmarkID): TBSBookmarkData; virtual;
    Function BookmarkRemove(ID: TBSBookmarkID): Integer; virtual;
    procedure BookmarkDelete(Index: Integer); virtual;
    procedure BookmarkClear; virtual;
    // bookmark information access
    Function BookmarkGetPosition(ID: TBSBookmarkID): Int64; virtual;
    procedure BookmarkSetPosition(ID: TBSBookmarkID; NewPosition: Int64); virtual;
    // position manipulation methods
    procedure MoveAt(Position: Int64); overload; virtual;
    procedure MoveAtOffset(Offset: Int64); virtual;
    procedure MoveBy(Delta: Int64); virtual;
    procedure MoveToStart; virtual;
    procedure MoveTo(ID: TBSBookmarkID); virtual;
    procedure MoveToIndex(Index: Integer); virtual;
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
    Function WriteSmallAnsiString(const Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUTF8String(const Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallWideString(const Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUnicodeString(const Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUCS4String(const Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallString(const Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBuffer(const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBytes(const Value: array of UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function FillBytes(Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteVariant(const Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function WriteBoolAt(Position: Int64; Value: ByteBool; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteBooleanAt(Position: Int64; Value: Boolean; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteInt8At(Position: Int64; Value: Int8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUInt8At(Position: Int64; Value: UInt8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteInt16At(Position: Int64; Value: Int16; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUInt16At(Position: Int64; Value: UInt16; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteInt32At(Position: Int64; Value: Int32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUInt32At(Position: Int64; Value: UInt32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteInt64At(Position: Int64; Value: Int64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUInt64At(Position: Int64; Value: UInt64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteFloat32At(Position: Int64; Value: Float32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteFloat64At(Position: Int64; Value: Float64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteFloat80At(Position: Int64; Value: Float80; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteDateTimeAt(Position: Int64; Value: TDateTime; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteCurrencyAt(Position: Int64; Value: Currency; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteAnsiCharAt(Position: Int64; Value: AnsiChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUTF8CharAt(Position: Int64; Value: UTF8Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteWideCharAt(Position: Int64; Value: WideChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUnicodeCharAt(Position: Int64; Value: UnicodeChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUCS4CharAt(Position: Int64; Value: UCS4Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteCharAt(Position: Int64; Value: Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteShortStringAt(Position: Int64; const Value: ShortString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteAnsiStringAt(Position: Int64; const Value: AnsiString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUTF8StringAt(Position: Int64; const Value: UTF8String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteWideStringAt(Position: Int64; const Value: WideString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUnicodeStringAt(Position: Int64; const Value: UnicodeString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUCS4StringAt(Position: Int64; const Value: UCS4String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteStringAt(Position: Int64; const Value: String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallAnsiStringAt(Position: Int64; const Value: AnsiString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallUTF8StringAt(Position: Int64; const Value: UTF8String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallWideStringAt(Position: Int64; const Value: WideString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallUnicodeStringAt(Position: Int64; const Value: UnicodeString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallUCS4StringAt(Position: Int64; const Value: UCS4String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallStringAt(Position: Int64; const Value: String; Advance: Boolean = True): TMemSize; overload; virtual;  
    Function WriteBufferAt(Position: Int64; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteBytesAt(Position: Int64; const Value: array of UInt8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function FillBytesAt(Position: Int64; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteVariantAt(Position: Int64; const Value: Variant; Advance: Boolean = True): TMemSize; overload; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function WriteBoolAtOffset(Offset: Int64; Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBooleanAtOffset(Offset: Int64; Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt8AtOffset(Offset: Int64; Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt8AtOffset(Offset: Int64; Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt16AtOffset(Offset: Int64; Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt16AtOffset(Offset: Int64; Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt32AtOffset(Offset: Int64; Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt32AtOffset(Offset: Int64; Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt64AtOffset(Offset: Int64; Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt64AtOffset(Offset: Int64; Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat32AtOffset(Offset: Int64; Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat64AtOffset(Offset: Int64; Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat80AtOffset(Offset: Int64; Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function WriteDateTimeAtOffset(Offset: Int64; Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function WriteCurrencyAtOffset(Offset: Int64; Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function WriteAnsiCharAtOffset(Offset: Int64; Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUTF8CharAtOffset(Offset: Int64; Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteWideCharAtOffset(Offset: Int64; Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUnicodeCharAtOffset(Offset: Int64; Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUCS4CharAtOffset(Offset: Int64; Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteCharAtOffset(Offset: Int64; Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteShortStringAtOffset(Offset: Int64; const Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteAnsiStringAtOffset(Offset: Int64; const Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUTF8StringAtOffset(Offset: Int64; const Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteWideStringAtOffset(Offset: Int64; const Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUnicodeStringAtOffset(Offset: Int64; const Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUCS4StringAtOffset(Offset: Int64; const Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteStringAtOffset(Offset: Int64; const Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallAnsiStringAtOffset(Offset: Int64; const Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUTF8StringAtOffset(Offset: Int64; const Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallWideStringAtOffset(Offset: Int64; const Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUnicodeStringAtOffset(Offset: Int64; const Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUCS4StringAtOffset(Offset: Int64; const Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallStringAtOffset(Offset: Int64; const Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBufferAtOffset(Offset: Int64; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBytesAtOffset(Offset: Int64; const Value: array of UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function FillBytesAtOffset(Offset: Int64; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteVariantAtOffset(Offset: Int64; const Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function WriteBoolTo(ID: TBSBookmarkID; Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBooleanTo(ID: TBSBookmarkID; Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt8To(ID: TBSBookmarkID; Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt8To(ID: TBSBookmarkID; Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt16To(ID: TBSBookmarkID; Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt16To(ID: TBSBookmarkID; Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt32To(ID: TBSBookmarkID; Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt32To(ID: TBSBookmarkID; Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt64To(ID: TBSBookmarkID; Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt64To(ID: TBSBookmarkID; Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat32To(ID: TBSBookmarkID; Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat64To(ID: TBSBookmarkID; Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat80To(ID: TBSBookmarkID; Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function WriteDateTimeTo(ID: TBSBookmarkID; Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function WriteCurrencyTo(ID: TBSBookmarkID; Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function WriteAnsiCharTo(ID: TBSBookmarkID; Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUTF8CharTo(ID: TBSBookmarkID; Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteWideCharTo(ID: TBSBookmarkID; Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUnicodeCharTo(ID: TBSBookmarkID; Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUCS4CharTo(ID: TBSBookmarkID; Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteCharTo(ID: TBSBookmarkID; Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteShortStringTo(ID: TBSBookmarkID; const Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteAnsiStringTo(ID: TBSBookmarkID; const Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUTF8StringTo(ID: TBSBookmarkID; const Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteWideStringTo(ID: TBSBookmarkID; const Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUnicodeStringTo(ID: TBSBookmarkID; const Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUCS4StringTo(ID: TBSBookmarkID; const Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteStringTo(ID: TBSBookmarkID; const Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallAnsiStringTo(ID: TBSBookmarkID; const Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUTF8StringTo(ID: TBSBookmarkID; const Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallWideStringTo(ID: TBSBookmarkID; const Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUnicodeStringTo(ID: TBSBookmarkID; const Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUCS4StringTo(ID: TBSBookmarkID; const Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallStringTo(ID: TBSBookmarkID; const Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBufferTo(ID: TBSBookmarkID; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBytesTo(ID: TBSBookmarkID; const Value: array of UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function FillBytesTo(ID: TBSBookmarkID; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteVariantTo(ID: TBSBookmarkID; const Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function WriteBoolToIndex(Index: Integer; Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBooleanToIndex(Index: Integer; Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt8ToIndex(Index: Integer; Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt8ToIndex(Index: Integer; Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt16ToIndex(Index: Integer; Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt16ToIndex(Index: Integer; Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt32ToIndex(Index: Integer; Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt32ToIndex(Index: Integer; Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteInt64ToIndex(Index: Integer; Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUInt64ToIndex(Index: Integer; Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat32ToIndex(Index: Integer; Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat64ToIndex(Index: Integer; Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function WriteFloat80ToIndex(Index: Integer; Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function WriteDateTimeToIndex(Index: Integer; Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function WriteCurrencyToIndex(Index: Integer; Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function WriteAnsiCharToIndex(Index: Integer; Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUTF8CharToIndex(Index: Integer; Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteWideCharToIndex(Index: Integer; Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUnicodeCharToIndex(Index: Integer; Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUCS4CharToIndex(Index: Integer; Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteCharToIndex(Index: Integer; Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function WriteShortStringToIndex(Index: Integer; const Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteAnsiStringToIndex(Index: Integer; const Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUTF8StringToIndex(Index: Integer; const Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteWideStringToIndex(Index: Integer; const Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUnicodeStringToIndex(Index: Integer; const Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteUCS4StringToIndex(Index: Integer; const Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteStringToIndex(Index: Integer; const Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallAnsiStringToIndex(Index: Integer; const Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUTF8StringToIndex(Index: Integer; const Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallWideStringToIndex(Index: Integer; const Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUnicodeStringToIndex(Index: Integer; const Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallUCS4StringToIndex(Index: Integer; const Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteSmallStringToIndex(Index: Integer; const Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBufferToIndex(Index: Integer; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function WriteBytesToIndex(Index: Integer; const Value: array of UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function FillBytesToIndex(Index: Integer; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function WriteVariantToIndex(Index: Integer; const Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // read methods
    Function ReadBool(out Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function ReadBoolean(out Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt8(out Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt8(out Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt16(out Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt16(out Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt32(out Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt32(out Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt64(out Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt64(out Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat32(out Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat64(out Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat80(out Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function ReadDateTime(out Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function ReadCurrency(out Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function ReadAnsiChar(out Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUTF8Char(out Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadWideChar(out Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUnicodeChar(out Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUCS4Char(out Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadChar(out Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadShortString(out Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadAnsiString(out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUTF8String(out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadWideString(out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUnicodeString(out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUCS4String(out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadString(out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallAnsiString(out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUTF8String(out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallWideString(out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUnicodeString(out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUCS4String(out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallString(out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadBuffer(out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function ReadVariant(out Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function ReadBoolAt(Position: Int64; out Value: ByteBool; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadBooleanAt(Position: Int64; out Value: Boolean; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadInt8At(Position: Int64; out Value: Int8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUInt8At(Position: Int64; out Value: UInt8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadInt16At(Position: Int64; out Value: Int16; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUInt16At(Position: Int64; out Value: UInt16; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadInt32At(Position: Int64; out Value: Int32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUInt32At(Position: Int64; out Value: UInt32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadInt64At(Position: Int64; out Value: Int64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUInt64At(Position: Int64; out Value: UInt64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadFloat32At(Position: Int64; out Value: Float32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadFloat64At(Position: Int64; out Value: Float64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadFloat80At(Position: Int64; out Value: Float80; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadDateTimeAt(Position: Int64; out Value: TDateTime; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadCurrencyAt(Position: Int64; out Value: Currency; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadAnsiCharAt(Position: Int64; out Value: AnsiChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUTF8CharAt(Position: Int64; out Value: UTF8Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadWideCharAt(Position: Int64; out Value: WideChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUnicodeCharAt(Position: Int64; out Value: UnicodeChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUCS4CharAt(Position: Int64; out Value: UCS4Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadCharAt(Position: Int64; out Value: Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadShortStringAt(Position: Int64; out Value: ShortString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadAnsiStringAt(Position: Int64; out Value: AnsiString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUTF8StringAt(Position: Int64; out Value: UTF8String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadWideStringAt(Position: Int64; out Value: WideString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUnicodeStringAt(Position: Int64; out Value: UnicodeString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUCS4StringAt(Position: Int64; out Value: UCS4String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadStringAt(Position: Int64; out Value: String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallAnsiStringAt(Position: Int64; out Value: AnsiString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallUTF8StringAt(Position: Int64; out Value: UTF8String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallWideStringAt(Position: Int64; out Value: WideString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallUnicodeStringAt(Position: Int64; out Value: UnicodeString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallUCS4StringAt(Position: Int64; out Value: UCS4String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallStringAt(Position: Int64; out Value: String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadBufferAt(Position: Int64; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadVariantAt(Position: Int64; out Value: Variant; Advance: Boolean = True): TMemSize; overload; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function ReadBoolAtOffset(Offset: Int64; out Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function ReadBooleanAtOffset(Offset: Int64; out Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt8AtOffset(Offset: Int64; out Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt8AtOffset(Offset: Int64; out Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt16AtOffset(Offset: Int64; out Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt16AtOffset(Offset: Int64; out Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt32AtOffset(Offset: Int64; out Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt32AtOffset(Offset: Int64; out Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt64AtOffset(Offset: Int64; out Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt64AtOffset(Offset: Int64; out Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat32AtOffset(Offset: Int64; out Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat64AtOffset(Offset: Int64; out Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat80AtOffset(Offset: Int64; out Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function ReadDateTimeAtOffset(Offset: Int64; out Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function ReadCurrencyAtOffset(Offset: Int64; out Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function ReadAnsiCharAtOffset(Offset: Int64; out Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUTF8CharAtOffset(Offset: Int64; out Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadWideCharAtOffset(Offset: Int64; out Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUnicodeCharAtOffset(Offset: Int64; out Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUCS4CharAtOffset(Offset: Int64; out Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadCharAtOffset(Offset: Int64; out Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadShortStringAtOffset(Offset: Int64; out Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadAnsiStringAtOffset(Offset: Int64; out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUTF8StringAtOffset(Offset: Int64; out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadWideStringAtOffset(Offset: Int64; out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUnicodeStringAtOffset(Offset: Int64; out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUCS4StringAtOffset(Offset: Int64; out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadStringAtOffset(Offset: Int64; out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallAnsiStringAtOffset(Offset: Int64; out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUTF8StringAtOffset(Offset: Int64; out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallWideStringAtOffset(Offset: Int64; out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUnicodeStringAtOffset(Offset: Int64; out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUCS4StringAtOffset(Offset: Int64; out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallStringAtOffset(Offset: Int64; out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadBufferAtOffset(Offset: Int64; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function ReadVariantAtOffset(Offset: Int64; out Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function ReadBoolFrom(ID: TBSBookmarkID; out Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function ReadBooleanFrom(ID: TBSBookmarkID; out Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt8From(ID: TBSBookmarkID; out Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt8From(ID: TBSBookmarkID; out Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt16From(ID: TBSBookmarkID; out Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt16From(ID: TBSBookmarkID; out Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt32From(ID: TBSBookmarkID; out Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt32From(ID: TBSBookmarkID; out Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt64From(ID: TBSBookmarkID; out Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt64From(ID: TBSBookmarkID; out Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat32From(ID: TBSBookmarkID; out Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat64From(ID: TBSBookmarkID; out Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat80From(ID: TBSBookmarkID; out Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function ReadDateTimeFrom(ID: TBSBookmarkID; out Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function ReadCurrencyFrom(ID: TBSBookmarkID; out Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function ReadAnsiCharFrom(ID: TBSBookmarkID; out Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUTF8CharFrom(ID: TBSBookmarkID; out Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadWideCharFrom(ID: TBSBookmarkID; out Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUnicodeCharFrom(ID: TBSBookmarkID; out Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUCS4CharFrom(ID: TBSBookmarkID; out Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadCharFrom(ID: TBSBookmarkID; out Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadShortStringFrom(ID: TBSBookmarkID; out Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadAnsiStringFrom(ID: TBSBookmarkID; out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUTF8StringFrom(ID: TBSBookmarkID; out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadWideStringFrom(ID: TBSBookmarkID; out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUnicodeStringFrom(ID: TBSBookmarkID; out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUCS4StringFrom(ID: TBSBookmarkID; out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadStringFrom(ID: TBSBookmarkID; out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallAnsiStringFrom(ID: TBSBookmarkID; out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUTF8StringFrom(ID: TBSBookmarkID; out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallWideStringFrom(ID: TBSBookmarkID; out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUnicodeStringFrom(ID: TBSBookmarkID; out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUCS4StringFrom(ID: TBSBookmarkID; out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallStringFrom(ID: TBSBookmarkID; out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadBufferFrom(ID: TBSBookmarkID; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function ReadVariantFrom(ID: TBSBookmarkID; out Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function ReadBoolFromIndex(Index: Integer; out Value: ByteBool; Advance: Boolean = True): TMemSize; virtual;
    Function ReadBooleanFromIndex(Index: Integer; out Value: Boolean; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt8FromIndex(Index: Integer; out Value: Int8; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt8FromIndex(Index: Integer; out Value: UInt8; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt16FromIndex(Index: Integer; out Value: Int16; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt16FromIndex(Index: Integer; out Value: UInt16; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt32FromIndex(Index: Integer; out Value: Int32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt32FromIndex(Index: Integer; out Value: UInt32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadInt64FromIndex(Index: Integer; out Value: Int64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUInt64FromIndex(Index: Integer; out Value: UInt64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat32FromIndex(Index: Integer; out Value: Float32; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat64FromIndex(Index: Integer; out Value: Float64; Advance: Boolean = True): TMemSize; virtual;
    Function ReadFloat80FromIndex(Index: Integer; out Value: Float80; Advance: Boolean = True): TMemSize; virtual;
    Function ReadDateTimeFromIndex(Index: Integer; out Value: TDateTime; Advance: Boolean = True): TMemSize; virtual;
    Function ReadCurrencyFromIndex(Index: Integer; out Value: Currency; Advance: Boolean = True): TMemSize; virtual;
    Function ReadAnsiCharFromIndex(Index: Integer; out Value: AnsiChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUTF8CharFromIndex(Index: Integer; out Value: UTF8Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadWideCharFromIndex(Index: Integer; out Value: WideChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUnicodeCharFromIndex(Index: Integer; out Value: UnicodeChar; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUCS4CharFromIndex(Index: Integer; out Value: UCS4Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadCharFromIndex(Index: Integer; out Value: Char; Advance: Boolean = True): TMemSize; virtual;
    Function ReadShortStringFromIndex(Index: Integer; out Value: ShortString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadAnsiStringFromIndex(Index: Integer; out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUTF8StringFromIndex(Index: Integer; out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadWideStringFromIndex(Index: Integer; out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUnicodeStringFromIndex(Index: Integer; out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadUCS4StringFromIndex(Index: Integer; out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadStringFromIndex(Index: Integer; out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallAnsiStringFromIndex(Index: Integer; out Value: AnsiString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUTF8StringFromIndex(Index: Integer; out Value: UTF8String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallWideStringFromIndex(Index: Integer; out Value: WideString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUnicodeStringFromIndex(Index: Integer; out Value: UnicodeString; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallUCS4StringFromIndex(Index: Integer; out Value: UCS4String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadSmallStringFromIndex(Index: Integer; out Value: String; Advance: Boolean = True): TMemSize; virtual;
    Function ReadBufferFromIndex(Index: Integer; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; virtual;
    Function ReadVariantFromIndex(Index: Integer; out Value: Variant; Advance: Boolean = True): TMemSize; virtual;
    // reading into result
    Function GetBool(Advance: Boolean = True): ByteBool; virtual;
    Function GetBoolean(Advance: Boolean = True): Boolean; virtual;
    Function GetInt8(Advance: Boolean = True): Int8; virtual;
    Function GetUInt8(Advance: Boolean = True): UInt8; virtual;
    Function GetInt16(Advance: Boolean = True): Int16; virtual;
    Function GetUInt16(Advance: Boolean = True): UInt16; virtual;
    Function GetInt32(Advance: Boolean = True): Int32; virtual;
    Function GetUInt32(Advance: Boolean = True): UInt32; virtual;
    Function GetInt64(Advance: Boolean = True): Int64; virtual;
    Function GetUInt64(Advance: Boolean = True): UInt64; virtual;
    Function GetFloat32(Advance: Boolean = True): Float32; virtual;
    Function GetFloat64(Advance: Boolean = True): Float64; virtual;
    Function GetFloat80(Advance: Boolean = True): Float80; virtual;
    Function GetDateTime(Advance: Boolean = True): TDateTime; virtual;
    Function GetCurrency(Advance: Boolean = True): Currency; virtual;
    Function GetAnsiChar(Advance: Boolean = True): AnsiChar; virtual;
    Function GetUTF8Char(Advance: Boolean = True): UTF8Char; virtual;
    Function GetWideChar(Advance: Boolean = True): WideChar; virtual;
    Function GetUnicodeChar(Advance: Boolean = True): UnicodeChar; virtual;
    Function GetUCS4Char(Advance: Boolean = True): UCS4Char; virtual;
    Function GetChar(Advance: Boolean = True): Char; virtual;
    Function GetShortString(Advance: Boolean = True): ShortString; virtual;
    Function GetAnsiString(Advance: Boolean = True): AnsiString; virtual;
    Function GetUTF8String(Advance: Boolean = True): UTF8String; virtual;
    Function GetWideString(Advance: Boolean = True): WideString; virtual;
    Function GetUnicodeString(Advance: Boolean = True): UnicodeString; virtual;
    Function GetUCS4String(Advance: Boolean = True): UCS4String; virtual;
    Function GetString(Advance: Boolean = True): String; virtual;
    Function GetSmallAnsiString(Advance: Boolean = True): AnsiString; virtual;
    Function GetSmallUTF8String(Advance: Boolean = True): UTF8String; virtual;
    Function GetSmallWideString(Advance: Boolean = True): WideString; virtual;
    Function GetSmallUnicodeString(Advance: Boolean = True): UnicodeString; virtual;
    Function GetSmallUCS4String(Advance: Boolean = True): UCS4String; virtual;
    Function GetSmallString(Advance: Boolean = True): String; virtual;
    Function GetVariant(Advance: Boolean = True): Variant; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function GetBoolAt(Position: Int64; Advance: Boolean = True): ByteBool; overload; virtual;
    Function GetBooleanAt(Position: Int64; Advance: Boolean = True): Boolean; overload; virtual;
    Function GetInt8At(Position: Int64; Advance: Boolean = True): Int8; overload; virtual;
    Function GetUInt8At(Position: Int64; Advance: Boolean = True): UInt8; overload; virtual;
    Function GetInt16At(Position: Int64; Advance: Boolean = True): Int16; overload; virtual;
    Function GetUInt16At(Position: Int64; Advance: Boolean = True): UInt16; overload; virtual;
    Function GetInt32At(Position: Int64; Advance: Boolean = True): Int32; overload; virtual;
    Function GetUInt32At(Position: Int64; Advance: Boolean = True): UInt32; overload; virtual;
    Function GetInt64At(Position: Int64; Advance: Boolean = True): Int64; overload; virtual;
    Function GetUInt64At(Position: Int64; Advance: Boolean = True): UInt64; overload; virtual;
    Function GetFloat32At(Position: Int64; Advance: Boolean = True): Float32; overload; virtual;
    Function GetFloat64At(Position: Int64; Advance: Boolean = True): Float64; overload; virtual;
    Function GetFloat80At(Position: Int64; Advance: Boolean = True): Float80; overload; virtual;
    Function GetDateTimeAt(Position: Int64; Advance: Boolean = True): TDateTime; overload; virtual;
    Function GetCurrencyAt(Position: Int64; Advance: Boolean = True): Currency; overload; virtual;
    Function GetAnsiCharAt(Position: Int64; Advance: Boolean = True): AnsiChar; overload; virtual;
    Function GetUTF8CharAt(Position: Int64; Advance: Boolean = True): UTF8Char; overload; virtual;
    Function GetWideCharAt(Position: Int64; Advance: Boolean = True): WideChar; overload; virtual;
    Function GetUnicodeCharAt(Position: Int64; Advance: Boolean = True): UnicodeChar; overload; virtual;
    Function GetUCS4CharAt(Position: Int64; Advance: Boolean = True): UCS4Char; overload; virtual;
    Function GetCharAt(Position: Int64; Advance: Boolean = True): Char; overload; virtual;
    Function GetShortStringAt(Position: Int64; Advance: Boolean = True): ShortString; overload; virtual;
    Function GetAnsiStringAt(Position: Int64; Advance: Boolean = True): AnsiString; overload; virtual;
    Function GetUTF8StringAt(Position: Int64; Advance: Boolean = True): UTF8String; overload; virtual;
    Function GetWideStringAt(Position: Int64; Advance: Boolean = True): WideString; overload; virtual;
    Function GetUnicodeStringAt(Position: Int64; Advance: Boolean = True): UnicodeString; overload; virtual;
    Function GetUCS4StringAt(Position: Int64; Advance: Boolean = True): UCS4String; overload; virtual;
    Function GetStringAt(Position: Int64; Advance: Boolean = True): String; overload; virtual;
    Function GetSmallAnsiStringAt(Position: Int64; Advance: Boolean = True): AnsiString; overload; virtual;
    Function GetSmallUTF8StringAt(Position: Int64; Advance: Boolean = True): UTF8String; overload; virtual;
    Function GetSmallWideStringAt(Position: Int64; Advance: Boolean = True): WideString; overload; virtual;
    Function GetSmallUnicodeStringAt(Position: Int64; Advance: Boolean = True): UnicodeString; overload; virtual;
    Function GetSmallUCS4StringAt(Position: Int64; Advance: Boolean = True): UCS4String; overload; virtual;
    Function GetSmallStringAt(Position: Int64; Advance: Boolean = True): String; overload; virtual;
    Function GetVariantAt(Position: Int64; Advance: Boolean = True): Variant; overload; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function GetBoolAtOffset(Offset: Int64; Advance: Boolean = True): ByteBool; virtual;
    Function GetBooleanAtOffset(Offset: Int64; Advance: Boolean = True): Boolean; virtual;
    Function GetInt8AtOffset(Offset: Int64; Advance: Boolean = True): Int8; virtual;
    Function GetUInt8AtOffset(Offset: Int64; Advance: Boolean = True): UInt8; virtual;
    Function GetInt16AtOffset(Offset: Int64; Advance: Boolean = True): Int16; virtual;
    Function GetUInt16AtOffset(Offset: Int64; Advance: Boolean = True): UInt16; virtual;
    Function GetInt32AtOffset(Offset: Int64; Advance: Boolean = True): Int32; virtual;
    Function GetUInt32AtOffset(Offset: Int64; Advance: Boolean = True): UInt32; virtual;
    Function GetInt64AtOffset(Offset: Int64; Advance: Boolean = True): Int64; virtual;
    Function GetUInt64AtOffset(Offset: Int64; Advance: Boolean = True): UInt64; virtual;
    Function GetFloat32AtOffset(Offset: Int64; Advance: Boolean = True): Float32; virtual;
    Function GetFloat64AtOffset(Offset: Int64; Advance: Boolean = True): Float64; virtual;
    Function GetFloat80AtOffset(Offset: Int64; Advance: Boolean = True): Float80; virtual;
    Function GetDateTimeAtOffset(Offset: Int64; Advance: Boolean = True): TDateTime; virtual;
    Function GetCurrencyAtOffset(Offset: Int64; Advance: Boolean = True): Currency; virtual;
    Function GetAnsiCharAtOffset(Offset: Int64; Advance: Boolean = True): AnsiChar; virtual;
    Function GetUTF8CharAtOffset(Offset: Int64; Advance: Boolean = True): UTF8Char; virtual;
    Function GetWideCharAtOffset(Offset: Int64; Advance: Boolean = True): WideChar; virtual;
    Function GetUnicodeCharAtOffset(Offset: Int64; Advance: Boolean = True): UnicodeChar; virtual;
    Function GetUCS4CharAtOffset(Offset: Int64; Advance: Boolean = True): UCS4Char; virtual;
    Function GetCharAtOffset(Offset: Int64; Advance: Boolean = True): Char; virtual;
    Function GetShortStringAtOffset(Offset: Int64; Advance: Boolean = True): ShortString; virtual;
    Function GetAnsiStringAtOffset(Offset: Int64; Advance: Boolean = True): AnsiString; virtual;
    Function GetUTF8StringAtOffset(Offset: Int64; Advance: Boolean = True): UTF8String; virtual;
    Function GetWideStringAtOffset(Offset: Int64; Advance: Boolean = True): WideString; virtual;
    Function GetUnicodeStringAtOffset(Offset: Int64; Advance: Boolean = True): UnicodeString; virtual;
    Function GetUCS4StringAtOffset(Offset: Int64; Advance: Boolean = True): UCS4String; virtual;
    Function GetStringAtOffset(Offset: Int64; Advance: Boolean = True): String; virtual;
    Function GetSmallAnsiStringAtOffset(Offset: Int64; Advance: Boolean = True): AnsiString; virtual;
    Function GetSmallUTF8StringAtOffset(Offset: Int64; Advance: Boolean = True): UTF8String; virtual;
    Function GetSmallWideStringAtOffset(Offset: Int64; Advance: Boolean = True): WideString; virtual;
    Function GetSmallUnicodeStringAtOffset(Offset: Int64; Advance: Boolean = True): UnicodeString; virtual;
    Function GetSmallUCS4StringAtOffset(Offset: Int64; Advance: Boolean = True): UCS4String; virtual;
    Function GetSmallStringAtOffset(Offset: Int64; Advance: Boolean = True): String; virtual;
    Function GetVariantAtOffset(Offset: Int64; Advance: Boolean = True): Variant; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function GetBoolFrom(ID: TBSBookmarkID; Advance: Boolean = True): ByteBool; virtual;
    Function GetBooleanFrom(ID: TBSBookmarkID; Advance: Boolean = True): Boolean; virtual;
    Function GetInt8From(ID: TBSBookmarkID; Advance: Boolean = True): Int8; virtual;
    Function GetUInt8From(ID: TBSBookmarkID; Advance: Boolean = True): UInt8; virtual;
    Function GetInt16From(ID: TBSBookmarkID; Advance: Boolean = True): Int16; virtual;
    Function GetUInt16From(ID: TBSBookmarkID; Advance: Boolean = True): UInt16; virtual;
    Function GetInt32From(ID: TBSBookmarkID; Advance: Boolean = True): Int32; virtual;
    Function GetUInt32From(ID: TBSBookmarkID; Advance: Boolean = True): UInt32; virtual;
    Function GetInt64From(ID: TBSBookmarkID; Advance: Boolean = True): Int64; virtual;
    Function GetUInt64From(ID: TBSBookmarkID; Advance: Boolean = True): UInt64; virtual;
    Function GetFloat32From(ID: TBSBookmarkID; Advance: Boolean = True): Float32; virtual;
    Function GetFloat64From(ID: TBSBookmarkID; Advance: Boolean = True): Float64; virtual;
    Function GetFloat80From(ID: TBSBookmarkID; Advance: Boolean = True): Float80; virtual;
    Function GetDateTimeFrom(ID: TBSBookmarkID; Advance: Boolean = True): TDateTime; virtual;
    Function GetCurrencyFrom(ID: TBSBookmarkID; Advance: Boolean = True): Currency; virtual;
    Function GetAnsiCharFrom(ID: TBSBookmarkID; Advance: Boolean = True): AnsiChar; virtual;
    Function GetUTF8CharFrom(ID: TBSBookmarkID; Advance: Boolean = True): UTF8Char; virtual;
    Function GetWideCharFrom(ID: TBSBookmarkID; Advance: Boolean = True): WideChar; virtual;
    Function GetUnicodeCharFrom(ID: TBSBookmarkID; Advance: Boolean = True): UnicodeChar; virtual;
    Function GetUCS4CharFrom(ID: TBSBookmarkID; Advance: Boolean = True): UCS4Char; virtual;
    Function GetCharFrom(ID: TBSBookmarkID; Advance: Boolean = True): Char; virtual;
    Function GetShortStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): ShortString; virtual;
    Function GetAnsiStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): AnsiString; virtual;
    Function GetUTF8StringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UTF8String; virtual;
    Function GetWideStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): WideString; virtual;
    Function GetUnicodeStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UnicodeString; virtual;
    Function GetUCS4StringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UCS4String; virtual;
    Function GetStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): String; virtual;
    Function GetSmallAnsiStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): AnsiString; virtual;
    Function GetSmallUTF8StringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UTF8String; virtual;
    Function GetSmallWideStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): WideString; virtual;
    Function GetSmallUnicodeStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UnicodeString; virtual;
    Function GetSmallUCS4StringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UCS4String; virtual;
    Function GetSmallStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): String; virtual;
    Function GetVariantFrom(ID: TBSBookmarkID; Advance: Boolean = True): Variant; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function GetBoolFromIndex(Index: Integer; Advance: Boolean = True): ByteBool; virtual;
    Function GetBooleanFromIndex(Index: Integer; Advance: Boolean = True): Boolean; virtual;
    Function GetInt8FromIndex(Index: Integer; Advance: Boolean = True): Int8; virtual;
    Function GetUInt8FromIndex(Index: Integer; Advance: Boolean = True): UInt8; virtual;
    Function GetInt16FromIndex(Index: Integer; Advance: Boolean = True): Int16; virtual;
    Function GetUInt16FromIndex(Index: Integer; Advance: Boolean = True): UInt16; virtual;
    Function GetInt32FromIndex(Index: Integer; Advance: Boolean = True): Int32; virtual;
    Function GetUInt32FromIndex(Index: Integer; Advance: Boolean = True): UInt32; virtual;
    Function GetInt64FromIndex(Index: Integer; Advance: Boolean = True): Int64; virtual;
    Function GetUInt64FromIndex(Index: Integer; Advance: Boolean = True): UInt64; virtual;
    Function GetFloat32FromIndex(Index: Integer; Advance: Boolean = True): Float32; virtual;
    Function GetFloat64FromIndex(Index: Integer; Advance: Boolean = True): Float64; virtual;
    Function GetFloat80FromIndex(Index: Integer; Advance: Boolean = True): Float80; virtual;
    Function GetDateTimeFromIndex(Index: Integer; Advance: Boolean = True): TDateTime; virtual;
    Function GetCurrencyFromIndex(Index: Integer; Advance: Boolean = True): Currency; virtual;
    Function GetAnsiCharFromIndex(Index: Integer; Advance: Boolean = True): AnsiChar; virtual;
    Function GetUTF8CharFromIndex(Index: Integer; Advance: Boolean = True): UTF8Char; virtual;
    Function GetWideCharFromIndex(Index: Integer; Advance: Boolean = True): WideChar; virtual;
    Function GetUnicodeCharFromIndex(Index: Integer; Advance: Boolean = True): UnicodeChar; virtual;
    Function GetUCS4CharFromIndex(Index: Integer; Advance: Boolean = True): UCS4Char; virtual;
    Function GetCharFromIndex(Index: Integer; Advance: Boolean = True): Char; virtual;
    Function GetShortStringFromIndex(Index: Integer; Advance: Boolean = True): ShortString; virtual;
    Function GetAnsiStringFromIndex(Index: Integer; Advance: Boolean = True): AnsiString; virtual;
    Function GetUTF8StringFromIndex(Index: Integer; Advance: Boolean = True): UTF8String; virtual;
    Function GetWideStringFromIndex(Index: Integer; Advance: Boolean = True): WideString; virtual;
    Function GetUnicodeStringFromIndex(Index: Integer; Advance: Boolean = True): UnicodeString; virtual;
    Function GetUCS4StringFromIndex(Index: Integer; Advance: Boolean = True): UCS4String; virtual;
    Function GetStringFromIndex(Index: Integer; Advance: Boolean = True): String; virtual;
    Function GetSmallAnsiStringFromIndex(Index: Integer; Advance: Boolean = True): AnsiString; virtual;
    Function GetSmallUTF8StringFromIndex(Index: Integer; Advance: Boolean = True): UTF8String; virtual;
    Function GetSmallWideStringFromIndex(Index: Integer; Advance: Boolean = True): WideString; virtual;
    Function GetSmallUnicodeStringFromIndex(Index: Integer; Advance: Boolean = True): UnicodeString; virtual;
    Function GetSmallUCS4StringFromIndex(Index: Integer; Advance: Boolean = True): UCS4String; virtual;
    Function GetSmallStringFromIndex(Index: Integer; Advance: Boolean = True): String; virtual;
    Function GetVariantFromIndex(Index: Integer; Advance: Boolean = True): Variant; virtual;
    // properties
    property Endian: TEndian read fEndian write fEndian;
    property Start: Int64 read fStart;
    property Position: Int64 read GetPosition write SetPosition;
    property Offset: Int64 read GetOffset write SetOffset;
    property Bookmarks[Index: Integer]: TBSBookmarkData read GetBookmark write SetBookmark; default;
    property BookmarkPtrs[Index: Integer]: PBSBookmarkData read GetBookmarkPtr;
    property BookmarkCount: Integer read GetCount;
    property OnChange: TNotifyEvent read fOnChangeEvent write fOnChangeEvent;
    property OnChangeEvent: TNotifyEvent read fOnChangeEvent write fOnChangeEvent;
    property OnChangeCallback: TNotifyCallback read fOnChangeCallback write fOnChangeCallback;    
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
    fPositionAddress: Pointer;
    Function GetStartAddress: Pointer; virtual;
    Function GetPosition: Int64; override;
    procedure SetPosition(NewPosition: Int64); override;
    Function WriteValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; override;
    Function WriteValueAtAddr(Address: Pointer; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    Function ReadValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; override;
    Function ReadValueAtAddr(Address: Pointer; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; virtual;
    procedure Initialize(Memory: Pointer); reintroduce; virtual;
  public
    class Function AddressToPosition(Address: Pointer): Int64; virtual;
    class Function PositionToAddress(Position: Int64): Pointer; virtual;
    constructor Create(Memory: Pointer); overload;
    procedure MoveAt(Address: Pointer); overload; virtual;
    Function BookmarkAdd(ID: TBSBookmarkID; Address: Pointer): Integer; overload; virtual;
    procedure BookmarkInsert(Index: Integer; ID: TBSBookmarkID; Address: Pointer); overload; virtual;
    Function BookmarkGetAddress(ID: TBSBookmarkID): Pointer; virtual;
    procedure BookmarkSetAddress(ID: TBSBookmarkID; NewAddress: Pointer); virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function WriteBoolAt(Address: Pointer; Value: ByteBool; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteBooleanAt(Address: Pointer; Value: Boolean; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteInt8At(Address: Pointer; Value: Int8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUInt8At(Address: Pointer; Value: UInt8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteInt16At(Address: Pointer; Value: Int16; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUInt16At(Address: Pointer; Value: UInt16; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteInt32At(Address: Pointer; Value: Int32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUInt32At(Address: Pointer; Value: UInt32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteInt64At(Address: Pointer; Value: Int64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUInt64At(Address: Pointer; Value: UInt64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteFloat32At(Address: Pointer; Value: Float32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteFloat64At(Address: Pointer; Value: Float64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteFloat80At(Address: Pointer; Value: Float80; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteDateTimeAt(Address: Pointer; Value: TDateTime; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteCurrencyAt(Address: Pointer; Value: Currency; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteAnsiCharAt(Address: Pointer; Value: AnsiChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUTF8CharAt(Address: Pointer; Value: UTF8Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteWideCharAt(Address: Pointer; Value: WideChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUnicodeCharAt(Address: Pointer; Value: UnicodeChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUCS4CharAt(Address: Pointer; Value: UCS4Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteCharAt(Address: Pointer; Value: Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteShortStringAt(Address: Pointer; const Value: ShortString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteAnsiStringAt(Address: Pointer; const Value: AnsiString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUTF8StringAt(Address: Pointer; const Value: UTF8String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteWideStringAt(Address: Pointer; const Value: WideString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUnicodeStringAt(Address: Pointer; const Value: UnicodeString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteUCS4StringAt(Address: Pointer; const Value: UCS4String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteStringAt(Address: Pointer; const Value: String; Advance: Boolean = True): TMemSize; overload; virtual; 
    Function WriteSmallAnsiStringAt(Address: Pointer; const Value: AnsiString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallUTF8StringAt(Address: Pointer; const Value: UTF8String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallWideStringAt(Address: Pointer; const Value: WideString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallUnicodeStringAt(Address: Pointer; const Value: UnicodeString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallUCS4StringAt(Address: Pointer; const Value: UCS4String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteSmallStringAt(Address: Pointer; const Value: String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteBufferAt(Address: Pointer; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteBytesAt(Address: Pointer; const Value: array of UInt8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function FillBytesAt(Address: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function WriteVariantAt(Address: Pointer; const Value: Variant; Advance: Boolean = True): TMemSize; overload; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function ReadBoolAt(Address: Pointer; out Value: ByteBool; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadBooleanAt(Address: Pointer; out Value: Boolean; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadInt8At(Address: Pointer; out Value: Int8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUInt8At(Address: Pointer; out Value: UInt8; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadInt16At(Address: Pointer; out Value: Int16; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUInt16At(Address: Pointer; out Value: UInt16; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadInt32At(Address: Pointer; out Value: Int32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUInt32At(Address: Pointer; out Value: UInt32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadInt64At(Address: Pointer; out Value: Int64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUInt64At(Address: Pointer; out Value: UInt64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadFloat32At(Address: Pointer; out Value: Float32; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadFloat64At(Address: Pointer; out Value: Float64; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadFloat80At(Address: Pointer; out Value: Float80; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadDateTimeAt(Address: Pointer; out Value: TDateTime; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadCurrencyAt(Address: Pointer; out Value: Currency; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadAnsiCharAt(Address: Pointer; out Value: AnsiChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUTF8CharAt(Address: Pointer; out Value: UTF8Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadWideCharAt(Address: Pointer; out Value: WideChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUnicodeCharAt(Address: Pointer; out Value: UnicodeChar; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUCS4CharAt(Address: Pointer; out Value: UCS4Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadCharAt(Address: Pointer; out Value: Char; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadShortStringAt(Address: Pointer; out Value: ShortString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadAnsiStringAt(Address: Pointer; out Value: AnsiString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUTF8StringAt(Address: Pointer; out Value: UTF8String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadWideStringAt(Address: Pointer; out Value: WideString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUnicodeStringAt(Address: Pointer; out Value: UnicodeString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadUCS4StringAt(Address: Pointer; out Value: UCS4String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadStringAt(Address: Pointer; out Value: String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallAnsiStringAt(Address: Pointer; out Value: AnsiString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallUTF8StringAt(Address: Pointer; out Value: UTF8String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallWideStringAt(Address: Pointer; out Value: WideString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallUnicodeStringAt(Address: Pointer; out Value: UnicodeString; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallUCS4StringAt(Address: Pointer; out Value: UCS4String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadSmallStringAt(Address: Pointer; out Value: String; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadBufferAt(Address: Pointer; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize; overload; virtual;
    Function ReadVariantAt(Address: Pointer; out Value: Variant; Advance: Boolean = True): TMemSize; overload; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Function GetBoolAt(Address: Pointer; Advance: Boolean = True): ByteBool; overload; virtual;
    Function GetBooleanAt(Address: Pointer; Advance: Boolean = True): Boolean; overload; virtual;
    Function GetInt8At(Address: Pointer; Advance: Boolean = True): Int8; overload; virtual;
    Function GetUInt8At(Address: Pointer; Advance: Boolean = True): UInt8; overload; virtual;
    Function GetInt16At(Address: Pointer; Advance: Boolean = True): Int16; overload; virtual;
    Function GetUInt16At(Address: Pointer; Advance: Boolean = True): UInt16; overload; virtual;
    Function GetInt32At(Address: Pointer; Advance: Boolean = True): Int32; overload; virtual;
    Function GetUInt32At(Address: Pointer; Advance: Boolean = True): UInt32; overload; virtual;
    Function GetInt64At(Address: Pointer; Advance: Boolean = True): Int64; overload; virtual;
    Function GetUInt64At(Address: Pointer; Advance: Boolean = True): UInt64; overload; virtual;
    Function GetFloat32At(Address: Pointer; Advance: Boolean = True): Float32; overload; virtual;
    Function GetFloat64At(Address: Pointer; Advance: Boolean = True): Float64; overload; virtual;
    Function GetFloat80At(Address: Pointer; Advance: Boolean = True): Float80; overload; virtual;
    Function GetDateTimeAt(Address: Pointer; Advance: Boolean = True): TDateTime; overload; virtual;
    Function GetCurrencyAt(Address: Pointer; Advance: Boolean = True): Currency; overload; virtual;
    Function GetAnsiCharAt(Address: Pointer; Advance: Boolean = True): AnsiChar; overload; virtual;
    Function GetUTF8CharAt(Address: Pointer; Advance: Boolean = True): UTF8Char; overload; virtual;
    Function GetWideCharAt(Address: Pointer; Advance: Boolean = True): WideChar; overload; virtual;
    Function GetUnicodeCharAt(Address: Pointer; Advance: Boolean = True): UnicodeChar; overload; virtual;
    Function GetUCS4CharAt(Address: Pointer; Advance: Boolean = True): UCS4Char; overload; virtual;
    Function GetCharAt(Address: Pointer; Advance: Boolean = True): Char; overload; virtual;
    Function GetShortStringAt(Address: Pointer; Advance: Boolean = True): ShortString; overload; virtual;
    Function GetAnsiStringAt(Address: Pointer; Advance: Boolean = True): AnsiString; overload; virtual;
    Function GetUTF8StringAt(Address: Pointer; Advance: Boolean = True): UTF8String; overload; virtual;
    Function GetWideStringAt(Address: Pointer; Advance: Boolean = True): WideString; overload; virtual;
    Function GetUnicodeStringAt(Address: Pointer; Advance: Boolean = True): UnicodeString; overload; virtual;
    Function GetUCS4StringAt(Address: Pointer; Advance: Boolean = True): UCS4String; overload; virtual;
    Function GetStringAt(Address: Pointer; Advance: Boolean = True): String; overload; virtual;
    Function GetSmallAnsiStringAt(Address: Pointer; Advance: Boolean = True): AnsiString; overload; virtual;
    Function GetSmallUTF8StringAt(Address: Pointer; Advance: Boolean = True): UTF8String; overload; virtual;
    Function GetSmallWideStringAt(Address: Pointer; Advance: Boolean = True): WideString; overload; virtual;
    Function GetSmallUnicodeStringAt(Address: Pointer; Advance: Boolean = True): UnicodeString; overload; virtual;
    Function GetSmallUCS4StringAt(Address: Pointer; Advance: Boolean = True): UCS4String; overload; virtual;
    Function GetSmallStringAt(Address: Pointer; Advance: Boolean = True): String; overload; virtual;
    Function GetVariantAt(Address: Pointer; Advance: Boolean = True): Variant; overload; virtual;
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    property StartAddress: Pointer read GetStartAddress;
    property PositionAddress: Pointer read fPositionAddress write fPositionAddress;
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
    Function GetPosition: Int64; override;
    procedure SetPosition(NewPosition: Int64); override;
    Function WriteValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; override;
    Function ReadValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize; override;
    procedure Initialize(Target: TStream); reintroduce; virtual;
  public
    constructor Create(Target: TStream);
    property Target: TStream read fTarget;
  end;

implementation

{<lite-implementation>}

uses
  Variants, Math
  {$IFDEF FPC}, StrRect{$ENDIF};

{<lite-begin>}
{$IFDEF FPC_DisableWarns}
  {$DEFINE FPCDWM}
  {$DEFINE W4055:={$WARN 4055 OFF}} // Conversion between ordinals and pointers is not portable
  {<lite-remove-multi 9>}
  {$DEFINE W5024:={$WARN 5024 OFF}} // Parameter "$1" not used
  {$DEFINE W5058:={$WARN 5058 OFF}} // Variable "$1" does not seem to be initialized
  {$PUSH}{$WARN 2005 OFF}           // Comment level $1 found
  {$IF Defined(FPC) and (FPC_FULLVERSION >= 30200)}
    {$DEFINE W6058:={$WARN 6058 OFF}} // Call to subroutine "$1" marked as inline is not inlined
  {$ELSE}
    {$DEFINE W6058:=}
  {$IFEND}
  {$POP}
{$ENDIF} 

{$IF SizeOf(UInt32) <> SizeOf(UCS4Char)}  // just to be sure
  {$MESSAGE ERROR 'Type size mismatch (UInt32 - UCS4Char).'}
{$IFEND}

{===============================================================================
    Internal variables
===============================================================================}
var
  OpenByteArrayIsPacked: Boolean = True;
  OpenByteArrayStride:   TMemSize = 1;

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

{<lite-replace " overload;" "">}
procedure ClampStringLength(var StrLength: Int32); overload;{$IFDEF CanInline} inline;{$ENDIF}
begin
If StrLength < 0 then
  StrLength := 0;
end;

{<lite-remove-multi 8>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

procedure ClampStringLength(var StrLength: Int16); overload;{$IFDEF CanInline} inline;{$ENDIF}
begin
If StrLength < 0 then
  StrLength := 0;
end;

//------------------------------------------------------------------------------
type
  TFloat80Overlay = packed array[0..9] of Byte;

{<lite-insert "\x7B$IFDEF ENDIAN_BIG\x7D">}
//==============================================================================

Function SwapEndian(Value: UInt16): UInt16; overload;{$IFDEF CanInline} inline;{$ENDIF}
begin
Result := UInt16(Value shl 8) or UInt16(Value shr 8);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function SwapEndian(Value: UInt32): UInt32; overload;
type
  Int32Rec = packed record
    LoWord: UInt16;
    HiWord: UInt16;
  end;
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
  Buffering might not be needed when operating on TCustomMemoryStream and its
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

{<lite-insert "\x7B$ENDIF\x7D">}
{<lite-end-ln>}

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

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function VarToBool(const Value: Variant): Boolean;
begin
Result := Value;
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function VarToInt64(const Value: Variant): Int64;
begin
Result := Value;
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function VarToUInt64(const Value: Variant): UInt64;
begin
Result := Value;
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

{===============================================================================
    Endianness - implementation
===============================================================================}

Function ProbeEndian: TEndian;
var
  TestVar: packed record
    case integer of
      0:  (TestArr:     array[0..3] of Byte);
      1:  (TestInt32:   Int32);
      2:  (TestFloat32: Float32)
  end;
begin
// init for integer test
TestVar.TestArr[0] := $01;
TestVar.TestArr[1] := $02;
TestVar.TestArr[2] := $03;
TestVar.TestArr[3] := $04;
// test integer
If TestVar.TestInt32 = $04030201 then
  begin
    // integer is little endian, test whether float is also little...
    TestVar.TestFloat32 := 0.1234;  // 0x3DFCB924
    If TestVar.TestInt32 = $3DFCB924 then
      Result := endLittle
    else
      raise EBSInvalidValue.Create('ProbeEndian: Unsupported byte order.');
  end
else If TestVar.TestInt32 = $01020304 then
  begin
    // integer is big endian, test float...
    TestVar.TestFloat32 := 0.1234;
    If TestVar.TestInt32 = $3DFCB924 {the order here is correct!} then
      Result := endBig
    else
      raise EBSInvalidValue.Create('ProbeEndian: Unsupported byte order.');
  end
else
  raise EBSInvalidValue.Create('ProbeEndian: Unsupported byte order.');
end;

//------------------------------------------------------------------------------

Function ResolveEndian(Endian: TEndian): TEndian;
begin
case Endian of
  endSystem:  Result := SysEndian;
  endDefault: Result := endLittle;
else
  Result := Endian;
end;
end;


{<lite-begin>}
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

Function StreamedSize_ShortString(const Value: ShortString): TMemSize;
begin
Result := StreamedSize_UInt8 + TMemSize(Length(Value));
end;     

//------------------------------------------------------------------------------

Function StreamedSize_AnsiString(const Value: AnsiString): TMemSize;
begin
Result := StreamedSize_Int32 + TMemSize(Length(Value) * SizeOf(AnsiChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_UTF8String(const Value: UTF8String): TMemSize;
begin
Result := StreamedSize_Int32 + TMemSize(Length(Value) * SizeOf(UTF8Char));
end;

//------------------------------------------------------------------------------

Function StreamedSize_WideString(const Value: WideString): TMemSize;
begin
Result := StreamedSize_Int32 + TMemSize(Length(Value) * SizeOf(WideChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_UnicodeString(const Value: UnicodeString): TMemSize;
begin
Result := StreamedSize_Int32 + TMemSize(Length(Value) * SizeOf(UnicodeChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_UCS4String(const Value: UCS4String): TMemSize;
begin
{
  Note that UCS4 string CAN (we must cope with situation it does not) contain
  an explicit terminating zero, which is not being streamed.
}
If Length(Value) > 0 then
  Result := StreamedSize_Int32 + TMemSize(IfThen(Value[High(Value)] = 0,Pred(Length(Value)),Length(Value)) * SizeOf(UCS4Char))
else
  Result := StreamedSize_Int32;
end;

//------------------------------------------------------------------------------

Function StreamedSize_String(const Value: String): TMemSize;
begin
Result := StreamedSize_UTF8String(StrToUTF8(Value));
end; 
{<lite-end-ln>}
//==============================================================================

Function StreamedSize_SmallAnsiString(const Value: AnsiString): TMemSize;
begin
Result := StreamedSize_Int16 + TMemSize(Length(Value) * SizeOf(AnsiChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_SmallUTF8String(const Value: UTF8String): TMemSize;
begin
Result := StreamedSize_Int16 + TMemSize(Length(Value) * SizeOf(UTF8Char));
end;

//------------------------------------------------------------------------------

Function StreamedSize_SmallWideString(const Value: WideString): TMemSize;
begin
Result := StreamedSize_Int16 + TMemSize(Length(Value) * SizeOf(WideChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_SmallUnicodeString(const Value: UnicodeString): TMemSize;
begin
Result := StreamedSize_Int16 + TMemSize(Length(Value) * SizeOf(UnicodeChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_SmallUCS4String(const Value: UCS4String): TMemSize;
begin
If Length(Value) > 0 then
  Result := StreamedSize_Int16 + TMemSize(IfThen(Value[High(Value)] = 0,Pred(Length(Value)),Length(Value)) * SizeOf(UCS4Char))
else
  Result := StreamedSize_Int16;
end;

//------------------------------------------------------------------------------

Function StreamedSize_SmallString(const Value: String): TMemSize;
begin
Result := StreamedSize_SmallUTF8String(StrToUTF8(Value));
end;

//==============================================================================

Function StreamedSize_TinyAnsiString(const Value: AnsiString): TMemSize;
begin
Result := StreamedSize_UInt8 + TMemSize(Length(Value) * SizeOf(AnsiChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_TinyUTF8String(const Value: UTF8String): TMemSize;
begin
Result := StreamedSize_UInt8 + TMemSize(Length(Value) * SizeOf(UTF8Char));
end;

//------------------------------------------------------------------------------

Function StreamedSize_TinyWideString(const Value: WideString): TMemSize;
begin
Result := StreamedSize_UInt8 + TMemSize(Length(Value) * SizeOf(WideChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_TinyUnicodeString(const Value: UnicodeString): TMemSize;
begin
Result := StreamedSize_UInt8 + TMemSize(Length(Value) * SizeOf(UnicodeChar));
end;

//------------------------------------------------------------------------------

Function StreamedSize_TinyUCS4String(const Value: UCS4String): TMemSize;
begin
If Length(Value) > 0 then
  Result := StreamedSize_UInt8 + TMemSize(IfThen(Value[High(Value)] = 0,Pred(Length(Value)),Length(Value)) * SizeOf(UCS4Char))
else
  Result := StreamedSize_UInt8;
end;

//------------------------------------------------------------------------------

Function StreamedSize_TinyString(const Value: String): TMemSize;
begin
Result := StreamedSize_TinyUTF8String(Value);
end;
{<lite-begin>}
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
{<lite-end-lns 2>}

//==============================================================================

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
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
        Indices := nil;
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
{$IFDEF FPCDWM}{$POP}{$ENDIF}

{<lite-begin>}
{===============================================================================
--------------------------------------------------------------------------------
                                 Memory writing
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

{<lite-replace _Ptr Ptr>}
Function _Ptr_WriteBool(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize;
begin
UInt8(Dest^) := BoolToNum(Value);
Result := SizeOf(UInt8);
AdvancePointer(Advance,Dest,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteBool_LE(var Dest: Pointer; Value: ByteBool; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBool(Dest,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_WriteBool_LE(Dest: Pointer; Value: ByteBool): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_WriteBool(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

{<lite-replace _Ptr Ptr>}
Function _Ptr_WriteInt8(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize;
begin
Int8(Dest^) := Value;
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteInt8_LE(var Dest: Pointer; Value: Int8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteInt8(Dest,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_WriteInt8_LE(Dest: Pointer; Value: Int8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_WriteInt8(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Ptr Ptr>}
Function _Ptr_WriteUInt8(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize;
begin
UInt8(Dest^) := Value;
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteUInt8_LE(var Dest: Pointer; Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteUInt8(Dest,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_WriteUInt8_LE(Dest: Pointer; Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_WriteUInt8(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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
 
{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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
  
{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

{<lite-replace _Ptr Ptr>}
Function _Ptr_WriteAnsiChar(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize;
begin
AnsiChar(Dest^) := Value;
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiChar_LE(var Dest: Pointer; Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteAnsiChar(Dest,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_WriteAnsiChar_LE(Dest: Pointer; Value: AnsiChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_WriteAnsiChar(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Ptr Ptr>}
Function _Ptr_WriteUTF8Char(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize;
begin
UTF8Char(Dest^) := Value;
Result := SizeOf(Value);
AdvancePointer(Advance,Dest,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8Char_LE(var Dest: Pointer; Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteUTF8Char(Dest,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
              
{<lite-replace _LE "">}            
Function Ptr_WriteUTF8Char_LE(Dest: Pointer; Value: UTF8Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_WriteUTF8Char(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}  
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">} 
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">} 
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">} 
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
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

{<lite-block-replace-begin _LE "">}
{<lite-replace _Ptr Ptr>}
Function _Ptr_WriteShortString(var Dest: Pointer; const Value: ShortString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteUInt8_LE(WorkPtr,UInt8(Length(Value)),True);
If Length(Value) > 0 then
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,(Addr(Value[1]))^,Length(Value),True));
If Advance then
  Dest := WorkPtr;
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteShortString_LE(var Dest: Pointer; const Value: ShortString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteShortString(Dest,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_WriteShortString_LE(Dest: Pointer; const Value: ShortString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_WriteShortString(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteShortString_BE(var Dest: Pointer; const Value: ShortString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteShortString(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteShortString_BE(Dest: Pointer; const Value: ShortString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteShortString(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteShortString(var Dest: Pointer; const Value: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteShortString_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteShortString_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteShortString(Dest: Pointer; const Value: ShortString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteShortString_BE(Ptr,Value,False)
else
  Result := Ptr_WriteShortString_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_WriteAnsiString_LE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_LE(WorkPtr,Length(Value),True);
If Length(Value) > 0 then
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiString_LE(Dest: Pointer; const Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteAnsiString_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiString_BE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_BE(WorkPtr,Length(Value),True);
If Length(Value) > 0 then
  Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiString_BE(Dest: Pointer; const Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteAnsiString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteAnsiString(var Dest: Pointer; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteAnsiString_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteAnsiString_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteAnsiString(Dest: Pointer; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteAnsiString_BE(Ptr,Value,False)
else
  Result := Ptr_WriteAnsiString_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_WriteUTF8String_LE(var Dest: Pointer; const Value: UTF8String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_LE(WorkPtr,Length(Value),True);
If Length(Value) > 0 then
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PUTF8Char(Value)^,Length(Value) * SizeOf(UTF8Char),True));
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
Function Ptr_WriteUTF8String_LE(Dest: Pointer; const Value: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUTF8String_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8String_BE(var Dest: Pointer; const Value: UTF8String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_BE(WorkPtr,Length(Value),True);
If Length(Value) > 0 then
  Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PUTF8Char(Value)^,Length(Value) * SizeOf(UTF8Char),True));
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
Function Ptr_WriteUTF8String_BE(Dest: Pointer; const Value: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUTF8String_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUTF8String(var Dest: Pointer; const Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUTF8String_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUTF8String_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
Function Ptr_WriteUTF8String(Dest: Pointer; const Value: UTF8String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUTF8String_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUTF8String_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_WriteWideString_LE(var Dest: Pointer; const Value: WideString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_LE(WorkPtr,Length(Value),True);
If Length(Value) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Value)),Length(Value)));
{$ELSE}
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PWideChar(Value)^,Length(Value) * SizeOf(WideChar),True));
{$ENDIF}
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideString_LE(Dest: Pointer; const Value: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteWideString_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteWideString_BE(var Dest: Pointer; const Value: WideString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_BE(WorkPtr,Length(Value),True);
If Length(Value) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PWideChar(Value)^,Length(Value) * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Value)),Length(Value)));
{$ENDIF}
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideString_BE(Dest: Pointer; const Value: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteWideString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteWideString(var Dest: Pointer; const Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteWideString_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteWideString_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteWideString(Dest: Pointer; const Value: WideString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteWideString_BE(Ptr,Value,False)
else
  Result := Ptr_WriteWideString_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_WriteUnicodeString_LE(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_LE(WorkPtr,Length(Value),True);
If Length(Value) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Value)),Length(Value)));
{$ELSE}
  Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PUnicodeChar(Value)^,Length(Value) * SizeOf(UnicodeChar),True));
{$ENDIF}
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeString_LE(Dest: Pointer; const Value: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUnicodeString_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeString_BE(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
WorkPtr := Dest;
Result := Ptr_WriteInt32_BE(WorkPtr,Length(Value),True);
If Length(Value) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PUnicodeChar(Value)^,Length(Value) * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Value)),Length(Value)));
{$ENDIF}
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeString_BE(Dest: Pointer; const Value: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUnicodeString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUnicodeString(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUnicodeString_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUnicodeString_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUnicodeString(Dest: Pointer; const Value: UnicodeString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUnicodeString_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUnicodeString_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_WriteUCS4String_LE(var Dest: Pointer; const Value: UCS4String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
  TrueLen:  TStrSize;
begin
WorkPtr := Dest;
If Length(Value) > 0 then
  begin
    If Value[High(Value)] = 0 then
      TrueLen := Pred(Length(Value))
    else
      TrueLen := Length(Value);
    If TrueLen > 0 then
      begin
        Result := Ptr_WriteInt32_LE(WorkPtr,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Ptr_WriteUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Value[Low(Value)])),TrueLen));
      {$ELSE}
        Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,Addr(Value[Low(Value)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ENDIF}
      end
    else Result := Ptr_WriteInt32_LE(WorkPtr,0,True);
  end
else Result := Ptr_WriteInt32_LE(WorkPtr,0,True);
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4String_LE(Dest: Pointer; const Value: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUCS4String_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteUCS4String_BE(var Dest: Pointer; const Value: UCS4String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
  TrueLen:  TStrSize;
begin
WorkPtr := Dest;
If Length(Value) > 0 then
  begin
    If Value[High(Value)] = 0 then
      TrueLen := Pred(Length(Value))
    else
      TrueLen := Length(Value);
    If TrueLen > 0 then
      begin
        Result := Ptr_WriteInt32_BE(WorkPtr,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,Addr(Value[Low(Value)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ELSE}
        Inc(Result,Ptr_WriteUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Value[Low(Value)])),TrueLen));
      {$ENDIF}
      end
    else Result := Ptr_WriteInt32_BE(WorkPtr,0,True);
  end
else Result := Ptr_WriteInt32_BE(WorkPtr,0,True);
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4String_BE(Dest: Pointer; const Value: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteUCS4String_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteUCS4String(var Dest: Pointer; const Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUCS4String_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteUCS4String_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteUCS4String(Dest: Pointer; const Value: UCS4String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteUCS4String_BE(Ptr,Value,False)
else
  Result := Ptr_WriteUCS4String_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_WriteString_LE(var Dest: Pointer; const Value: String; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteUTF8String_LE(Dest,StrToUTF8(Value),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteString_LE(Dest: Pointer; const Value: String): TMemSize;
begin
Result := Ptr_WriteUTF8String_LE(Dest,StrToUTF8(Value));
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteString_BE(var Dest: Pointer; const Value: String; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteUTF8String_BE(Dest,StrToUTF8(Value),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteString_BE(Dest: Pointer; const Value: String): TMemSize;
begin
Result := Ptr_WriteUTF8String_BE(Dest,StrToUTF8(Value));
end;

//------------------------------------------------------------------------------

Function Ptr_WriteString(var Dest: Pointer; const Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteUTF8String(Dest,StrToUTF8(Value),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteString(Dest: Pointer; const Value: String; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteUTF8String(Dest,StrToUTF8(Value),Endian);
end;

{-------------------------------------------------------------------------------
    Small strings
-------------------------------------------------------------------------------}

Function Ptr_WriteSmallAnsiString_LE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(Int16) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteInt16_LE(WorkPtr,Int16(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallAnsiString_LE: String is too long (%d) for small storage.',[Length(Value)]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallAnsiString_LE(Dest: Pointer; const Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallAnsiString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallAnsiString_BE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(Int16) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteInt16_BE(WorkPtr,Int16(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallAnsiString_BE: String is too long (%d) for small storage.',[Length(Value)]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallAnsiString_BE(Dest: Pointer; const Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallAnsiString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallAnsiString(var Dest: Pointer; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallAnsiString_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteSmallAnsiString_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallAnsiString(Dest: Pointer; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallAnsiString_BE(Ptr,Value,False)
else
  Result := Ptr_WriteSmallAnsiString_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteSmallUTF8String_LE(var Dest: Pointer; const Value: UTF8String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(Int16) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteInt16_LE(WorkPtr,Int16(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PUTF8Char(Value)^,Length(Value) * SizeOf(UTF8Char),True));
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallUTF8String_LE: String is too long (%d) for small storage.',[Length(Value)]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUTF8String_LE(Dest: Pointer; const Value: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallUTF8String_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUTF8String_BE(var Dest: Pointer; const Value: UTF8String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(Int16) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteInt16_BE(WorkPtr,Int16(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PUTF8Char(Value)^,Length(Value) * SizeOf(UTF8Char),True));
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallUTF8String_BE: String is too long (%d) for small storage.',[Length(Value)]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUTF8String_BE(Dest: Pointer; const Value: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallUTF8String_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUTF8String(var Dest: Pointer; const Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallUTF8String_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteSmallUTF8String_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUTF8String(Dest: Pointer; const Value: UTF8String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallUTF8String_BE(Ptr,Value,False)
else
  Result := Ptr_WriteSmallUTF8String_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteSmallWideString_LE(var Dest: Pointer; const Value: WideString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(Int16) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteInt16_LE(WorkPtr,Int16(Length(Value)),True);
    If Length(Value) > 0 then
    {$IFDEF ENDIAN_BIG}
      Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Value)),Length(Value)));
    {$ELSE}
      Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PWideChar(Value)^,Length(Value) * SizeOf(WideChar),True));
    {$ENDIF}
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallWideString_LE: String is too long (%d) for small storage.',[Length(Value)]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallWideString_LE(Dest: Pointer; const Value: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallWideString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallWideString_BE(var Dest: Pointer; const Value: WideString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(Int16) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteInt16_BE(WorkPtr,Int16(Length(Value)),True);
    If Length(Value) > 0 then
    {$IFDEF ENDIAN_BIG}
      Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PWideChar(Value)^,Length(Value) * SizeOf(WideChar),True));
    {$ELSE}
      Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Value)),Length(Value)));
    {$ENDIF}
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallWideString_BE: String is too long (%d) for small storage.',[Length(Value)]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallWideString_BE(Dest: Pointer; const Value: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallWideString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallWideString(var Dest: Pointer; const Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallWideString_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteSmallWideString_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallWideString(Dest: Pointer; const Value: WideString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallWideString_BE(Ptr,Value,False)
else
  Result := Ptr_WriteSmallWideString_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteSmallUnicodeString_LE(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(Int16) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteInt16_LE(WorkPtr,Int16(Length(Value)),True);
    If Length(Value) > 0 then
    {$IFDEF ENDIAN_BIG}
      Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Value)),Length(Value)));
    {$ELSE}
      Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PUnicodeChar(Value)^,Length(Value) * SizeOf(UnicodeChar),True));
    {$ENDIF}
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallUnicodeString_LE: String is too long (%d) for small storage.',[Length(Value)]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUnicodeString_LE(Dest: Pointer; const Value: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallUnicodeString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUnicodeString_BE(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(Int16) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteInt16_BE(WorkPtr,Int16(Length(Value)),True);
    If Length(Value) > 0 then
    {$IFDEF ENDIAN_BIG}
      Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,PUnicodeChar(Value)^,Length(Value) * SizeOf(UnicodeChar),True));
    {$ELSE}
      Inc(Result,Ptr_WriteUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Value)),Length(Value)));
    {$ENDIF}
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallUnicodeString_BE: String is too long (%d) for small storage.',[Length(Value)]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUnicodeString_BE(Dest: Pointer; const Value: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallUnicodeString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUnicodeString(var Dest: Pointer; const Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallUnicodeString_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteSmallUnicodeString_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUnicodeString(Dest: Pointer; const Value: UnicodeString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallUnicodeString_BE(Ptr,Value,False)
else
  Result := Ptr_WriteSmallUnicodeString_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteSmallUCS4String_LE(var Dest: Pointer; const Value: UCS4String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
  TrueLen:  TStrSize;
begin
WorkPtr := Dest;
If Length(Value) > 0 then
  begin
    If Value[High(Value)] = 0 then
      TrueLen := Pred(Length(Value))
    else
      TrueLen := Length(Value);
    If TrueLen > High(Int16) then
      raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallUCS4String_LE: String is too long (%d) for small storage.',[Length(Value)])
    else If TrueLen > 0 then
      begin
        Result := Ptr_WriteInt16_LE(WorkPtr,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Ptr_WriteUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Value[Low(Value)])),TrueLen));
      {$ELSE}
        Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,Addr(Value[Low(Value)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ENDIF}
      end
    else Result := Ptr_WriteInt16_LE(WorkPtr,0,True);
  end
else Result := Ptr_WriteInt16_LE(WorkPtr,0,True);
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUCS4String_LE(Dest: Pointer; const Value: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallUCS4String_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUCS4String_BE(var Dest: Pointer; const Value: UCS4String; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
  TrueLen:  TStrSize;
begin
WorkPtr := Dest;
If Length(Value) > 0 then
  begin
    If Value[High(Value)] = 0 then
      TrueLen := Pred(Length(Value))
    else
      TrueLen := Length(Value);
    If TrueLen > High(Int16) then
      raise EBSStringTooLong.CreateFmt('Ptr_WriteSmallUCS4String_BE: String is too long (%d) for small storage.',[Length(Value)])
    else If TrueLen > 0 then
      begin
        Result := Ptr_WriteInt16_BE(WorkPtr,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Ptr_WriteBuffer_BE(WorkPtr,Addr(Value[Low(Value)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ELSE}
        Inc(Result,Ptr_WriteUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Value[Low(Value)])),TrueLen));
      {$ENDIF}
      end
    else Result := Ptr_WriteInt16_BE(WorkPtr,0,True);
  end
else Result := Ptr_WriteInt16_BE(WorkPtr,0,True);
If Advance then
  Dest := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUCS4String_BE(Dest: Pointer; const Value: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteSmallUCS4String_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallUCS4String(var Dest: Pointer; const Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallUCS4String_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteSmallUCS4String_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallUCS4String(Dest: Pointer; const Value: UCS4String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteSmallUCS4String_BE(Ptr,Value,False)
else
  Result := Ptr_WriteSmallUCS4String_LE(Ptr,Value,False);
end;

//==============================================================================

Function Ptr_WriteSmallString_LE(var Dest: Pointer; const Value: String; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteSmallUTF8String_LE(Dest,StrToUTF8(Value),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallString_LE(Dest: Pointer; const Value: String): TMemSize;
begin
Result := Ptr_WriteSmallUTF8String_LE(Dest,StrToUTF8(Value));
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallString_BE(var Dest: Pointer; const Value: String; Advance: Boolean): TMemSize;
begin
Result := Ptr_WriteSmallUTF8String_BE(Dest,StrToUTF8(Value),Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallString_BE(Dest: Pointer; const Value: String): TMemSize;
begin
Result := Ptr_WriteSmallUTF8String_BE(Dest,StrToUTF8(Value));
end;

//------------------------------------------------------------------------------

Function Ptr_WriteSmallString(var Dest: Pointer; const Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteSmallUTF8String(Dest,StrToUTF8(Value),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteSmallString(Dest: Pointer; const Value: String; Endian: TEndian = endDefault): TMemSize;
begin
Result := Ptr_WriteSmallUTF8String(Dest,StrToUTF8(Value),Endian);
end;

{-------------------------------------------------------------------------------
    Tiny strings
-------------------------------------------------------------------------------}

Function _Ptr_WriteTinyAnsiString(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize;
var
  WorkPtr:  Pointer;
begin
If Length(Value) <= High(UInt8) then
  begin
    WorkPtr := Dest;
    Result := Ptr_WriteUInt8_LE(WorkPtr,UInt8(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Ptr_WriteBuffer_LE(WorkPtr,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
    If Advance then
      Dest := WorkPtr;
  end
else raise EBSStringTooLong.CreateFmt('_Ptr_WriteTinyAnsiString: String is too long (%d) for tiny storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteTinyAnsiString_LE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteTinyAnsiString(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteTinyAnsiString_LE(Dest: Pointer; const Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteTinyAnsiString(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteTinyAnsiString_BE(var Dest: Pointer; const Value: AnsiString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteShortString(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteTinyAnsiString_BE(Dest: Pointer; const Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := _Ptr_WriteTinyAnsiString(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_WriteTinyAnsiString(var Dest: Pointer; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteTinyAnsiString_BE(Dest,Value,Advance)
else
  Result := Ptr_WriteTinyAnsiString_LE(Dest,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteTinyAnsiString(Dest: Pointer; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_WriteTinyAnsiString_BE(Ptr,Value,False)
else
  Result := Ptr_WriteTinyAnsiString_LE(Ptr,Value,False);
end;

//==============================================================================

{<lite-begin>}
{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

{<lite-replace _Ptr Ptr>}
Function _Ptr_WriteBuffer(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Move(Buffer,Dest^,Size);
Result := Size;
AdvancePointer(Advance,Dest,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteBuffer_LE(var Dest: Pointer; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBuffer(Dest,Buffer,Size,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
{<lite-replace _LE "">}
Function Ptr_WriteBuffer_LE(Dest: Pointer; const Buffer; Size: TMemSize): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_WriteBuffer(Ptr,Buffer,Size,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Ptr Ptr>}
Function _Ptr_WriteBytes(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize;
var
  WorkPtr:  PUInt8;
  i:        Integer;
begin
If OpenByteArrayIsPacked then
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
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_WriteBytes_LE(var Dest: Pointer; const Value: array of UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_WriteBytes(Dest,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  
{<lite-replace _LE "">}
Function Ptr_WriteBytes_LE(Dest: Pointer; const Value: array of UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_WriteBytes(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Ptr Ptr>}
Function _Ptr_FillBytes(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize;
begin
FillChar(Dest^,Count,Value);
Result := Count;
AdvancePointer(Advance,Dest,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_FillBytes_LE(var Dest: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_FillBytes(Dest,Count,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_FillBytes_LE(Dest: Pointer; Count: TMemSize; Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_FillBytes(Ptr,Count,Value,False);
end;
{<lite-end-lns 2>}

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

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function Ptr_WriteVariant_LE(var Dest: Pointer; const Value: Variant; Advance: Boolean): TMemSize;
{$DEFINE BS_INC_VW}{$DEFINE BS_INC_M}{$DEFINE BS_INC_L}
  {$INCLUDE '.\BinaryStreaming_var.inc'}
{$UNDEF BS_INC_VW}{$UNDEF BS_INC_M}{$UNDEF BS_INC_L}
{$IFDEF FPCDWM}{$POP}{$ENDIF}

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_WriteVariant_LE(Dest: Pointer; const Value: Variant): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Dest;
Result := Ptr_WriteVariant_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function Ptr_WriteVariant_BE(var Dest: Pointer; const Value: Variant; Advance: Boolean): TMemSize;
{$DEFINE BS_INC_VW}{$DEFINE BS_INC_M}
  {$INCLUDE '.\BinaryStreaming_var.inc'}
{$UNDEF BS_INC_VW}{$UNDEF BS_INC_M}
{$IFDEF FPCDWM}{$POP}{$ENDIF}

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


{<lite-begin>}
{===============================================================================
--------------------------------------------------------------------------------
                                 Memory reading
--------------------------------------------------------------------------------
===============================================================================}

{<lite-replace _Ptr Ptr>}
Function _Ptr_ReadBool(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize;
begin
Value := NumToBool(UInt8(Src^));
Result := SizeOf(UInt8);
AdvancePointer(Advance,Src,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadBool_LE(var Src: Pointer; out Value: ByteBool; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadBool(Src,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_ReadBool_LE(Src: Pointer; out Value: ByteBool): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_ReadBool(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-block-replace-begin _Ptr Ptr>}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

{<lite-replace _Ptr Ptr>}
Function _Ptr_ReadInt8(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize;
begin
Value := Int8(Src^);
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadInt8_LE(var Src: Pointer; out Value: Int8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadInt8(Src,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
{<lite-replace _LE "">}        
Function Ptr_ReadInt8_LE(Src: Pointer; out Value: Int8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_ReadInt8(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-block-replace-begin _Ptr Ptr>}
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
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Ptr Ptr>}
Function _Ptr_ReadUInt8(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize;
begin
Value := UInt8(Src^);
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadUInt8_LE(var Src: Pointer; out Value: UInt8; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadUInt8(Src,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_ReadUInt8_LE(Src: Pointer; out Value: UInt8): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_ReadUInt8(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-block-replace-begin _Ptr Ptr>}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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
 
 {<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

{<lite-replace _Ptr Ptr>}
Function _Ptr_ReadAnsiChar(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Value := AnsiChar(Src^);
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadAnsiChar_LE(var Src: Pointer; out Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadAnsiChar(Src,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_ReadAnsiChar_LE(Src: Pointer; out Value: AnsiChar): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_ReadAnsiChar(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-block-replace-begin _Ptr Ptr>}
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
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Ptr Ptr>}
Function _Ptr_ReadUTF8Char(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Value := UTF8Char(Src^);
Result := SizeOf(Value);
AdvancePointer(Advance,Src,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8Char_LE(var Src: Pointer; out Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadUTF8Char(Src,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_ReadUTF8Char_LE(Src: Pointer; out Value: UTF8Char): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_ReadUTF8Char(Ptr,Value,False);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-block-replace-begin _Ptr Ptr>}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

{<lite-replace _Ptr Ptr>}
Function _Ptr_ReadShortString(var Src: Pointer; out Value: ShortString; Advance: Boolean): TMemSize;
var
  StrLength:  UInt8;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadUInt8_LE(WorkPtr,StrLength,True);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,Addr(Value[1])^,StrLength,True));
If Advance then
  Src := WorkPtr;
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadShortString_LE(var Src: Pointer; out Value: ShortString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadShortString(Src,Value,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_ReadShortString_LE(Src: Pointer; out Value: ShortString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_ReadShortString(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadShortString_BE(var Src: Pointer; out Value: ShortString; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadShortString(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadShortString_BE(Src: Pointer; out Value: ShortString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := _Ptr_ReadShortString(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadShortString(var Src: Pointer; out Value: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadShortString_BE(Src,Value,Advance)
else
  Result := Ptr_ReadShortString_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadShortString(Src: Pointer; out Value: ShortString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadShortString_BE(Ptr,Value,False)
else
  Result := Ptr_ReadShortString_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
{<lite-block-replace-begin _Ptr Ptr>}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_ReadAnsiString_LE(var Src: Pointer; out Value: AnsiString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PAnsiChar(Value)^,StrLength * SizeOf(AnsiChar),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiString_LE(Src: Pointer; out Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadAnsiString_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadAnsiString_BE(var Src: Pointer; out Value: AnsiString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PAnsiChar(Value)^,StrLength * SizeOf(AnsiChar),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiString_BE(Src: Pointer; out Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadAnsiString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadAnsiString(var Src: Pointer; out Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadAnsiString_BE(Src,Value,Advance)
else
  Result := Ptr_ReadAnsiString_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadAnsiString(Src: Pointer; out Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadAnsiString_BE(Ptr,Value,False)
else
  Result := Ptr_ReadAnsiString_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_ReadUTF8String_LE(var Src: Pointer; out Value: UTF8String; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PUTF8Char(Value)^,StrLength * SizeOf(UTF8Char),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8String_LE(Src: Pointer; out Value: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUTF8String_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8String_BE(var Src: Pointer; out Value: UTF8String; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PUTF8Char(Value)^,StrLength * SizeOf(UTF8Char),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8String_BE(Src: Pointer; out Value: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUTF8String_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUTF8String(var Src: Pointer; out Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUTF8String_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUTF8String_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUTF8String(Src: Pointer; out Value: UTF8String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUTF8String_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUTF8String_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_ReadWideString_LE(var Src: Pointer; out Value: WideString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Value)),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PWideChar(Value)^,StrLength * SizeOf(WideChar),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadWideString_LE(Src: Pointer; out Value: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadWideString_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadWideString_BE(var Src: Pointer; out Value: WideString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PWideChar(Value)^,StrLength * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Value)),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadWideString_BE(Src: Pointer; out Value: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadWideString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadWideString(var Src: Pointer; out Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadWideString_BE(Src,Value,Advance)
else
  Result := Ptr_ReadWideString_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadWideString(Src: Pointer; out Value: WideString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadWideString_BE(Ptr,Value,False)
else
  Result := Ptr_ReadWideString_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_ReadUnicodeString_LE(var Src: Pointer; out Value: UnicodeString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Value)),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PUnicodeChar(Value)^,StrLength * SizeOf(UnicodeChar),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUnicodeString_LE(Src: Pointer; out Value: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUnicodeString_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeString_BE(var Src: Pointer; out Value: UnicodeString; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PUnicodeChar(Value)^,StrLength * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Value)),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUnicodeString_BE(Src: Pointer; out Value: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUnicodeString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUnicodeString(var Src: Pointer; out Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUnicodeString_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUnicodeString_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUnicodeString(Src: Pointer; out Value: UnicodeString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUnicodeString_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUnicodeString_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_ReadUCS4String_LE(var Src: Pointer; out Value: UCS4String; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := nil;
SetLength(Value,StrLength + 1);
Value[High(Value)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Value[Low(Value)])),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,Addr(Value[Low(Value)])^,StrLength * SizeOf(UCS4Char),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4String_LE(Src: Pointer; out Value: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUCS4String_LE(Ptr,Value,False);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadUCS4String_BE(var Src: Pointer; out Value: UCS4String; Advance: Boolean): TMemSize;
var
  StrLength:  Int32;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt32_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := nil;
SetLength(Value,StrLength + 1);
Value[High(Value)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,Addr(Value[Low(Value)])^,StrLength * SizeOf(UCS4Char),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Value[Low(Value)])),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4String_BE(Src: Pointer; out Value: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadUCS4String_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadUCS4String(var Src: Pointer; out Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUCS4String_BE(Src,Value,Advance)
else
  Result := Ptr_ReadUCS4String_LE(Src,Value,Advance)
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadUCS4String(Src: Pointer; out Value: UCS4String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadUCS4String_BE(Ptr,Value,False)
else
  Result := Ptr_ReadUCS4String_LE(Ptr,Value,False);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
//==============================================================================

Function Ptr_ReadString_LE(var Src: Pointer; out Value: String; Advance: Boolean): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String_LE(Src,TempStr,Advance);
Value := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadString_LE(Src: Pointer; out Value: String): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String_LE(Src,TempStr);
Value := UTF8ToStr(TempStr);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadString_BE(var Src: Pointer; out Value: String; Advance: Boolean): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String_BE(Src,TempStr,Advance);
Value := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadString_BE(Src: Pointer; out Value: String): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String_BE(Src,TempStr);
Value := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadString(var Src: Pointer; out Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String(Src,TempStr,Advance,Endian);
Value := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadString(Src: Pointer; out Value: String; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadUTF8String(Src,TempStr,Endian);
Value := UTF8ToStr(TempStr);
end;

{<lite-begin>}
{<lite-block-replace-begin _LE "">}
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
{<lite-end-ln>}

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
    Small strings
-------------------------------------------------------------------------------}

Function Ptr_ReadSmallAnsiString_LE(var Src: Pointer; out Value: AnsiString; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PAnsiChar(Value)^,StrLength * SizeOf(AnsiChar),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallAnsiString_LE(Src: Pointer; out Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallAnsiString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallAnsiString_BE(var Src: Pointer; out Value: AnsiString; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PAnsiChar(Value)^,StrLength * SizeOf(AnsiChar),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallAnsiString_BE(Src: Pointer; out Value: AnsiString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallAnsiString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallAnsiString(var Src: Pointer; out Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallAnsiString_BE(Src,Value,Advance)
else
  Result := Ptr_ReadSmallAnsiString_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallAnsiString(Src: Pointer; out Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallAnsiString_BE(Ptr,Value,False)
else
  Result := Ptr_ReadSmallAnsiString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallAnsiString_LE(var Src: Pointer; Advance: Boolean): AnsiString;
begin
Ptr_ReadSmallAnsiString_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallAnsiString_LE(Src: Pointer): AnsiString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallAnsiString_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallAnsiString_BE(var Src: Pointer; Advance: Boolean): AnsiString;
begin
Ptr_ReadSmallAnsiString_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallAnsiString_BE(Src: Pointer): AnsiString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallAnsiString_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallAnsiString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): AnsiString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallAnsiString_BE(Src,Advance)
else
  Result := Ptr_GetSmallAnsiString_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallAnsiString(Src: Pointer; Endian: TEndian = endDefault): AnsiString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallAnsiString_BE(Ptr,False)
else
  Result := Ptr_GetSmallAnsiString_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadSmallUTF8String_LE(var Src: Pointer; out Value: UTF8String; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PUTF8Char(Value)^,StrLength * SizeOf(UTF8Char),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUTF8String_LE(Src: Pointer; out Value: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallUTF8String_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUTF8String_BE(var Src: Pointer; out Value: UTF8String; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PUTF8Char(Value)^,StrLength * SizeOf(UTF8Char),True));
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUTF8String_BE(Src: Pointer; out Value: UTF8String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallUTF8String_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUTF8String(var Src: Pointer; out Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallUTF8String_BE(Src,Value,Advance)
else
  Result := Ptr_ReadSmallUTF8String_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUTF8String(Src: Pointer; out Value: UTF8String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallUTF8String_BE(Ptr,Value,False)
else
  Result := Ptr_ReadSmallUTF8String_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUTF8String_LE(var Src: Pointer; Advance: Boolean): UTF8String;
begin
Ptr_ReadSmallUTF8String_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUTF8String_LE(Src: Pointer): UTF8String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallUTF8String_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUTF8String_BE(var Src: Pointer; Advance: Boolean): UTF8String;
begin
Ptr_ReadSmallUTF8String_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUTF8String_BE(Src: Pointer): UTF8String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallUTF8String_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUTF8String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UTF8String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallUTF8String_BE(Src,Advance)
else
  Result := Ptr_GetSmallUTF8String_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUTF8String(Src: Pointer; Endian: TEndian = endDefault): UTF8String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallUTF8String_BE(Ptr,False)
else
  Result := Ptr_GetSmallUTF8String_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadSmallWideString_LE(var Src: Pointer; out Value: WideString; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Value)),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PWideChar(Value)^,StrLength * SizeOf(WideChar),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallWideString_LE(Src: Pointer; out Value: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallWideString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallWideString_BE(var Src: Pointer; out Value: WideString; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PWideChar(Value)^,StrLength * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PWideChar(Value)),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallWideString_BE(Src: Pointer; out Value: WideString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallWideString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallWideString(var Src: Pointer; out Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallWideString_BE(Src,Value,Advance)
else
  Result := Ptr_ReadSmallWideString_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallWideString(Src: Pointer; out Value: WideString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallWideString_BE(Ptr,Value,False)
else
  Result := Ptr_ReadSmallWideString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallWideString_LE(var Src: Pointer; Advance: Boolean): WideString;
begin
Ptr_ReadSmallWideString_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallWideString_LE(Src: Pointer): WideString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallWideString_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallWideString_BE(var Src: Pointer; Advance: Boolean): WideString;
begin
Ptr_ReadSmallWideString_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallWideString_BE(Src: Pointer): WideString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallWideString_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallWideString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): WideString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallWideString_BE(Src,Advance)
else
  Result := Ptr_GetSmallWideString_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallWideString(Src: Pointer; Endian: TEndian = endDefault): WideString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallWideString_BE(Ptr,False)
else
  Result := Ptr_GetSmallWideString_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadSmallUnicodeString_LE(var Src: Pointer; out Value: UnicodeString; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Value)),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,PUnicodeChar(Value)^,StrLength * SizeOf(UnicodeChar),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUnicodeString_LE(Src: Pointer; out Value: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallUnicodeString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUnicodeString_BE(var Src: Pointer; out Value: UnicodeString; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,PUnicodeChar(Value)^,StrLength * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt16Arr_SwapEndian(PUInt16(WorkPtr),PUInt16(PUnicodeChar(Value)),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUnicodeString_BE(Src: Pointer; out Value: UnicodeString): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallUnicodeString_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUnicodeString(var Src: Pointer; out Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallUnicodeString_BE(Src,Value,Advance)
else
  Result := Ptr_ReadSmallUnicodeString_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUnicodeString(Src: Pointer; out Value: UnicodeString; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallUnicodeString_BE(Ptr,Value,False)
else
  Result := Ptr_ReadSmallUnicodeString_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUnicodeString_LE(var Src: Pointer; Advance: Boolean): UnicodeString;
begin
Ptr_ReadSmallUnicodeString_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUnicodeString_LE(Src: Pointer): UnicodeString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallUnicodeString_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUnicodeString_BE(var Src: Pointer; Advance: Boolean): UnicodeString;
begin
Ptr_ReadSmallUnicodeString_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUnicodeString_BE(Src: Pointer): UnicodeString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallUnicodeString_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUnicodeString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallUnicodeString_BE(Src,Advance)
else
  Result := Ptr_GetSmallUnicodeString_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUnicodeString(Src: Pointer; Endian: TEndian = endDefault): UnicodeString;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallUnicodeString_BE(Ptr,False)
else
  Result := Ptr_GetSmallUnicodeString_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadSmallUCS4String_LE(var Src: Pointer; out Value: UCS4String; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_LE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := nil;
SetLength(Value,StrLength + 1);
Value[High(Value)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Value[Low(Value)])),StrLength));
{$ELSE}
  Inc(Result,Ptr_ReadBuffer_LE(WorkPtr,Addr(Value[Low(Value)])^,StrLength * SizeOf(UCS4Char),True));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUCS4String_LE(Src: Pointer; out Value: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallUCS4String_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUCS4String_BE(var Src: Pointer; out Value: UCS4String; Advance: Boolean): TMemSize;
var
  StrLength:  Int16;
  WorkPtr:    Pointer;
begin
WorkPtr := Src;
Result := Ptr_ReadInt16_BE(WorkPtr,StrLength,True);
ClampStringLength(StrLength);
Value := nil;
SetLength(Value,StrLength + 1);
Value[High(Value)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Ptr_ReadBuffer_BE(WorkPtr,Addr(Value[Low(Value)])^,StrLength * SizeOf(UCS4Char),True));
{$ELSE}
  Inc(Result,Ptr_ReadUInt32Arr_SwapEndian(PUInt32(WorkPtr),PUInt32(Addr(Value[Low(Value)])),StrLength));
{$ENDIF}
If Advance then
  Src := WorkPtr;
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUCS4String_BE(Src: Pointer; out Value: UCS4String): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadSmallUCS4String_BE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallUCS4String(var Src: Pointer; out Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallUCS4String_BE(Src,Value,Advance)
else
  Result := Ptr_ReadSmallUCS4String_LE(Src,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallUCS4String(Src: Pointer; out Value: UCS4String; Endian: TEndian = endDefault): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_ReadSmallUCS4String_BE(Ptr,Value,False)
else
  Result := Ptr_ReadSmallUCS4String_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUCS4String_LE(var Src: Pointer; Advance: Boolean): UCS4String;
begin
Ptr_ReadSmallUCS4String_LE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUCS4String_LE(Src: Pointer): UCS4String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallUCS4String_LE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUCS4String_BE(var Src: Pointer; Advance: Boolean): UCS4String;
begin
Ptr_ReadSmallUCS4String_BE(Src,Result,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUCS4String_BE(Src: Pointer): UCS4String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Ptr_ReadSmallUCS4String_BE(Ptr,Result,False);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallUCS4String(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): UCS4String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallUCS4String_BE(Src,Advance)
else
  Result := Ptr_GetSmallUCS4String_LE(Src,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallUCS4String(Src: Pointer; Endian: TEndian = endDefault): UCS4String;
var
  Ptr:  Pointer;
begin
Ptr := Src;
If ResolveEndian(Endian) = endBig then
  Result := Ptr_GetSmallUCS4String_BE(Ptr,False)
else
  Result := Ptr_GetSmallUCS4String_LE(Ptr,False);
end;

//==============================================================================

Function Ptr_ReadSmallString_LE(var Src: Pointer; out Value: String; Advance: Boolean): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadSmallUTF8String_LE(Src,TempStr,Advance);
Value := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallString_LE(Src: Pointer; out Value: String): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadSmallUTF8String_LE(Src,TempStr);
Value := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallString_BE(var Src: Pointer; out Value: String; Advance: Boolean): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadSmallUTF8String_BE(Src,TempStr,Advance);
Value := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallString_BE(Src: Pointer; out Value: String): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadSmallUTF8String_BE(Src,TempStr);
Value := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Ptr_ReadSmallString(var Src: Pointer; out Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadSmallUTF8String(Src,TempStr,Advance,Endian);
Value := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadSmallString(Src: Pointer; out Value: String; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Ptr_ReadSmallUTF8String(Src,TempStr,Endian);
Value := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallString_LE(var Src: Pointer; Advance: Boolean): String;
begin
Result := UTF8ToStr(Ptr_GetSmallUTF8String_LE(Src,Advance));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallString_LE(Src: Pointer): String;
begin
Result := UTF8ToStr(Ptr_GetSmallUTF8String_LE(Src));
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallString_BE(var Src: Pointer; Advance: Boolean): String;
begin
Result := UTF8ToStr(Ptr_GetSmallUTF8String_BE(Src,Advance));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallString_BE(Src: Pointer): String;
begin
Result := UTF8ToStr(Ptr_GetSmallUTF8String_BE(Src));
end;

//------------------------------------------------------------------------------

Function Ptr_GetSmallString(var Src: Pointer; Advance: Boolean; Endian: TEndian = endDefault): String;
begin
Result := UTF8ToStr(Ptr_GetSmallUTF8String(Src,Advance,Endian));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_GetSmallString(Src: Pointer; Endian: TEndian = endDefault): String;
begin
Result := UTF8ToStr(Ptr_GetSmallUTF8String(Src,Endian));
end;

{<lite-begin>}
{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

{<lite-replace _Ptr Ptr>}
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
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Ptr_ReadBuffer_LE(var Src: Pointer; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Result := _Ptr_ReadBuffer(Src,Buffer,Size,Advance);
end;

{<lite-begin>}
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{<lite-replace _LE "">}
Function Ptr_ReadBuffer_LE(Src: Pointer; out Buffer; Size: TMemSize): TMemSize;
var
  Ptr:  Pointer;
begin
Ptr := Src;
{<lite-replace _Ptr Ptr>}
Result := _Ptr_ReadBuffer(Ptr,Buffer,Size,False);
end;
{<lite-end-lns 2>}

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

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function Ptr_ReadVariant_LE(var Src: Pointer; out Value: Variant; Advance: Boolean): TMemSize;
{$DEFINE BS_INC_VR}{$DEFINE BS_INC_M}{$DEFINE BS_INC_L}
  {$INCLUDE '.\BinaryStreaming_var.inc'}
{$UNDEF BS_INC_VR}{$UNDEF BS_INC_M}{$UNDEF BS_INC_L}
{$IFDEF FPCDWM}{$POP}{$ENDIF}

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Ptr_ReadVariant_LE(Src: Pointer; out Value: Variant): TMemSize; overload;
var
  Ptr:  Pointer;
begin
Ptr := Src;
Result := Ptr_ReadVariant_LE(Ptr,Value,False);
end;

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function Ptr_ReadVariant_BE(var Src: Pointer; out Value: Variant; Advance: Boolean): TMemSize;
{$DEFINE BS_INC_VR}{$DEFINE BS_INC_M}
  {$INCLUDE '.\BinaryStreaming_var.inc'}
{$UNDEF BS_INC_VR}{$UNDEF BS_INC_M}
{$IFDEF FPCDWM}{$POP}{$ENDIF}

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


{<lite-begin>}
{===============================================================================
--------------------------------------------------------------------------------
                                 Stream writing
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

{<lite-replace _Stream Stream>}
Function _Stream_WriteBool(Stream: TStream; Value: ByteBool; Advance: Boolean): TMemSize;
var
  Temp: UInt8;
begin
Temp := BoolToNum(Value);
Stream.WriteBuffer(Temp,SizeOf(Temp));
Result := SizeOf(Temp);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteBoolean_LE(Stream: TStream; Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteBool_LE(Stream,Value,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

{<lite-replace _Stream Stream>}
Function _Stream_WriteInt8(Stream: TStream; Value: Int8; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Stream Stream>}
Function _Stream_WriteUInt8(Stream: TStream; Value: UInt8; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_WriteInt16_LE(Stream: TStream; Value: Int16; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int16(SwapEndian(UInt16(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>} 
//==============================================================================

{<lite-replace _LE "">}
Function Stream_WriteUInt16_LE(Stream: TStream; Value: UInt16; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>} 
//==============================================================================

{<lite-replace _LE "">}
Function Stream_WriteInt32_LE(Stream: TStream; Value: Int32; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int32(SwapEndian(UInt32(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>} 
//==============================================================================

{<lite-replace _LE "">}
Function Stream_WriteUInt32_LE(Stream: TStream; Value: UInt32; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>} 
//==============================================================================

{<lite-replace _LE "">}
Function Stream_WriteInt64_LE(Stream: TStream; Value: Int64; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := Int64(SwapEndian(UInt64(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>} 
//==============================================================================

{<lite-replace _LE "">}
Function Stream_WriteUInt64_LE(Stream: TStream; Value: UInt64; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>} 
{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================
 
{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================
 
{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteDateTime_LE(Stream: TStream; Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteFloat64_LE(Stream,Value,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

{<lite-replace _Stream Stream>}
Function _Stream_WriteAnsiChar(Stream: TStream; Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Stream Stream>}
Function _Stream_WriteUTF8Char(Stream: TStream; Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_WriteWideChar_LE(Stream: TStream; Value: WideChar; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := WideChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_WriteUnicodeChar_LE(Stream: TStream; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
{$IFDEF ENDIAN_BIG}
Value := UnicodeChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Stream.WriteBuffer(Value,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteChar_LE(Stream: TStream; Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteUInt16_LE(Stream,UInt16(Ord(Value)),Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

{<lite-block-replace-begin _LE "">}
{<lite-replace _Stream Stream>}
Function _Stream_WriteShortString(Stream: TStream; const Value: ShortString; Advance: Boolean): TMemSize;
begin
Result := Stream_WriteUInt8_LE(Stream,UInt8(Length(Value)),True);
If Length(Value) > 0 then
  Inc(Result,Stream_WriteBuffer_LE(Stream,Addr(Value[1])^,Length(Value),True));
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_WriteShortString_LE(Stream: TStream; const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteShortString(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteShortString_BE(Stream: TStream; const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_WriteShortString(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteShortString(Stream: TStream; const Value: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteShortString_BE(Stream,Value,Advance)
else
  Result := Stream_WriteShortString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteShortString(Stream: TStream; const Value: ShortString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteShortString_BE(Stream,Value)
else
  Result := Stream_WriteShortString_LE(Stream,Value);
end;

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteAnsiString_LE(Stream: TStream; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_LE(Stream,Length(Value),True);
If Length(Value) > 0 then
  Inc(Result,Stream_WriteBuffer_LE(Stream,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_WriteAnsiString_BE(Stream: TStream; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_BE(Stream,Length(Value),True);
If Length(Value) > 0 then
  Inc(Result,Stream_WriteBuffer_BE(Stream,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteAnsiString(Stream: TStream; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteAnsiString_BE(Stream,Value,Advance)
else
  Result := Stream_WriteAnsiString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteAnsiString(Stream: TStream; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteAnsiString_BE(Stream,Value)
else
  Result := Stream_WriteAnsiString_LE(Stream,Value);
end;

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteUTF8String_LE(Stream: TStream; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_LE(Stream,Length(Value),True);
If Length(Value) > 0 then
  Inc(Result,Stream_WriteBuffer_LE(Stream,PUTF8Char(Value)^,Length(Value) * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_WriteUTF8String_BE(Stream: TStream; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_BE(Stream,Length(Value),True);
If Length(Value) > 0 then
  Inc(Result,Stream_WriteBuffer_BE(Stream,PUTF8Char(Value)^,Length(Value) * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUTF8String(Stream: TStream; const Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUTF8String_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUTF8String_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUTF8String(Stream: TStream; const Value: UTF8String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUTF8String_BE(Stream,Value)
else
  Result := Stream_WriteUTF8String_LE(Stream,Value);
end;

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteWideString_LE(Stream: TStream; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_LE(Stream,Length(Value),True);
If Length(Value) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Value)),Length(Value)));
{$ELSE}
  Inc(Result,Stream_WriteBuffer_LE(Stream,PWideChar(Value)^,Length(Value) * SizeOf(WideChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_WriteWideString_BE(Stream: TStream; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_BE(Stream,Length(Value),True);
If Length(Value) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_WriteBuffer_BE(Stream,PWideChar(Value)^,Length(Value) * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Value)),Length(Value)));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteWideString(Stream: TStream; const Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteWideString_BE(Stream,Value,Advance)
else
  Result := Stream_WriteWideString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteWideString(Stream: TStream; const Value: WideString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteWideString_BE(Stream,Value)
else
  Result := Stream_WriteWideString_LE(Stream,Value);
end;

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteUnicodeString_LE(Stream: TStream; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_LE(Stream,Length(Value),True);
If Length(Value) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Value)),Length(Value)));
{$ELSE}
  Inc(Result,Stream_WriteBuffer_LE(Stream,PUnicodeChar(Value)^,Length(Value) * SizeOf(UnicodeChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_WriteUnicodeString_BE(Stream: TStream; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteInt32_BE(Stream,Length(Value),True);
If Length(Value) > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_WriteBuffer_BE(Stream,PUnicodeChar(Value)^,Length(Value) * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Value)),Length(Value)));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUnicodeString(Stream: TStream; const Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUnicodeString_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUnicodeString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUnicodeString(Stream: TStream; const Value: UnicodeString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUnicodeString_BE(Stream,Value)
else
  Result := Stream_WriteUnicodeString_LE(Stream,Value);
end;

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteUCS4String_LE(Stream: TStream; const Value: UCS4String; Advance: Boolean = True): TMemSize;
var
  TrueLen:  TStrSize;
begin
If Length(Value) > 0 then
  begin
    If Value[High(Value)] = 0 then
      TrueLen := Pred(Length(Value))
    else
      TrueLen := Length(Value);
    If TrueLen > 0 then
      begin
        Result := Stream_WriteInt32_LE(Stream,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Stream_WriteUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Value[Low(Value)])),TrueLen));
      {$ELSE}
        Inc(Result,Stream_WriteBuffer_LE(Stream,Addr(Value[Low(Value)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ENDIF}
      end
    else Result := Stream_WriteInt32_LE(Stream,0,True);
  end
else Result := Stream_WriteInt32_LE(Stream,0,True);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_WriteUCS4String_BE(Stream: TStream; const Value: UCS4String; Advance: Boolean = True): TMemSize;
var
  TrueLen:  TStrSize;
begin
If Length(Value) > 0 then
  begin
    If Value[High(Value)] = 0 then
      TrueLen := Pred(Length(Value))
    else
      TrueLen := Length(Value);
    If TrueLen > 0 then
      begin
        Result := Stream_WriteInt32_BE(Stream,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Stream_WriteBuffer_BE(Stream,Addr(Value[Low(Value)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ELSE}
        Inc(Result,Stream_WriteUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Value[Low(Value)])),TrueLen));
      {$ENDIF}
      end
    else Result := Stream_WriteInt32_BE(Stream,0,True);
  end
else Result := Stream_WriteInt32_BE(Stream,0,True);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteUCS4String(Stream: TStream; const Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUCS4String_BE(Stream,Value,Advance)
else
  Result := Stream_WriteUCS4String_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteUCS4String(Stream: TStream; const Value: UCS4String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteUCS4String_BE(Stream,Value)
else
  Result := Stream_WriteUCS4String_LE(Stream,Value);
end;

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_WriteString_LE(Stream: TStream; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteUTF8String_LE(Stream,StrToUTF8(Value),Advance);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_WriteString_BE(Stream: TStream; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteUTF8String_BE(Stream,StrToUTF8(Value),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteString(Stream: TStream; const Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteUTF8String(Stream,StrToUTF8(Value),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteString(Stream: TStream; const Value: String; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteUTF8String(Stream,StrToUTF8(Value),Endian);
end;

{-------------------------------------------------------------------------------
    Small strings
-------------------------------------------------------------------------------}

Function Stream_WriteSmallAnsiString_LE(Stream: TStream; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
If Length(Value) <= High(Int16) then
  begin
    Result := Stream_WriteInt16_LE(Stream,Int16(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Stream_WriteBuffer_LE(Stream,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
    AdvanceStream(Advance,Stream,Result);
  end
else raise EBSStringTooLong.CreateFmt('Stream_WriteSmallAnsiString_LE: String is too long (%d) for small storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallAnsiString_BE(Stream: TStream; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
If Length(Value) <= High(Int16) then
  begin
    Result := Stream_WriteInt16_BE(Stream,Int16(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Stream_WriteBuffer_BE(Stream,PAnsiChar(Value)^,Length(Value) * SizeOf(AnsiChar),True));
    AdvanceStream(Advance,Stream,Result);
  end
else raise EBSStringTooLong.CreateFmt('Stream_WriteSmallAnsiString_BE: String is too long (%d) for small storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallAnsiString(Stream: TStream; const Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallAnsiString_BE(Stream,Value,Advance)
else
  Result := Stream_WriteSmallAnsiString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteSmallAnsiString(Stream: TStream; const Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallAnsiString_BE(Stream,Value)
else
  Result := Stream_WriteSmallAnsiString_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteSmallUTF8String_LE(Stream: TStream; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
If Length(Value) <= High(Int16) then
  begin
    Result := Stream_WriteInt16_LE(Stream,Int16(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Stream_WriteBuffer_LE(Stream,PUTF8Char(Value)^,Length(Value) * SizeOf(UTF8Char),True));
    AdvanceStream(Advance,Stream,Result);
  end
else raise EBSStringTooLong.CreateFmt('Stream_WriteSmallUTF8String_LE: String is too long (%d) for small storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUTF8String_BE(Stream: TStream; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
If Length(Value) <= High(Int16) then
  begin
    Result := Stream_WriteInt16_BE(Stream,Int16(Length(Value)),True);
    If Length(Value) > 0 then
      Inc(Result,Stream_WriteBuffer_BE(Stream,PUTF8Char(Value)^,Length(Value) * SizeOf(UTF8Char),True));
    AdvanceStream(Advance,Stream,Result);
  end
else raise EBSStringTooLong.CreateFmt('Stream_WriteSmallUTF8String_BE: String is too long (%d) for small storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUTF8String(Stream: TStream; const Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallUTF8String_BE(Stream,Value,Advance)
else
  Result := Stream_WriteSmallUTF8String_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteSmallUTF8String(Stream: TStream; const Value: UTF8String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallUTF8String_BE(Stream,Value)
else
  Result := Stream_WriteSmallUTF8String_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteSmallWideString_LE(Stream: TStream; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
If Length(Value) <= High(Int16) then
  begin
    Result := Stream_WriteInt16_LE(Stream,Int16(Length(Value)),True);
    If Length(Value) > 0 then
    {$IFDEF ENDIAN_BIG}
      Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Value)),Length(Value)));
    {$ELSE}
      Inc(Result,Stream_WriteBuffer_LE(Stream,PWideChar(Value)^,Length(Value) * SizeOf(WideChar),True));
    {$ENDIF}
    AdvanceStream(Advance,Stream,Result);
  end
else raise EBSStringTooLong.CreateFmt('Stream_WriteSmallWideString_LE: String is too long (%d) for small storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallWideString_BE(Stream: TStream; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
If Length(Value) <= High(Int16) then
  begin
    Result := Stream_WriteInt16_BE(Stream,Int16(Length(Value)),True);
    If Length(Value) > 0 then
    {$IFDEF ENDIAN_BIG}
      Inc(Result,Stream_WriteBuffer_BE(Stream,PWideChar(Value)^,Length(Value) * SizeOf(WideChar),True));
    {$ELSE}
      Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Value)),Length(Value)));
    {$ENDIF}
    AdvanceStream(Advance,Stream,Result);
  end
else raise EBSStringTooLong.CreateFmt('Stream_WriteSmallWideString_BE: String is too long (%d) for small storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallWideString(Stream: TStream; const Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallWideString_BE(Stream,Value,Advance)
else
  Result := Stream_WriteSmallWideString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteSmallWideString(Stream: TStream; const Value: WideString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallWideString_BE(Stream,Value)
else
  Result := Stream_WriteSmallWideString_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteSmallUnicodeString_LE(Stream: TStream; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
If Length(Value) <= High(Int16) then
  begin
    Result := Stream_WriteInt16_LE(Stream,Int16(Length(Value)),True);
    If Length(Value) > 0 then
    {$IFDEF ENDIAN_BIG}
      Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Value)),Length(Value)));
    {$ELSE}
      Inc(Result,Stream_WriteBuffer_LE(Stream,PUnicodeChar(Value)^,Length(Value) * SizeOf(UnicodeChar),True));
    {$ENDIF}
    AdvanceStream(Advance,Stream,Result);
  end
else raise EBSStringTooLong.CreateFmt('Stream_WriteSmallUnicodeString_LE: String is too long (%d) for small storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUnicodeString_BE(Stream: TStream; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
If Length(Value) <= High(Int16) then
  begin
    Result := Stream_WriteInt16_BE(Stream,Int16(Length(Value)),True);
    If Length(Value) > 0 then
    {$IFDEF ENDIAN_BIG}
      Inc(Result,Stream_WriteBuffer_BE(Stream,PUnicodeChar(Value)^,Length(Value) * SizeOf(UnicodeChar),True));
    {$ELSE}
      Inc(Result,Stream_WriteUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Value)),Length(Value)));
    {$ENDIF}
    AdvanceStream(Advance,Stream,Result);
  end
else raise EBSStringTooLong.CreateFmt('Stream_WriteSmallUnicodeString_BE: String is too long (%d) for small storage.',[Length(Value)]);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUnicodeString(Stream: TStream; const Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallUnicodeString_BE(Stream,Value,Advance)
else
  Result := Stream_WriteSmallUnicodeString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteSmallUnicodeString(Stream: TStream; const Value: UnicodeString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallUnicodeString_BE(Stream,Value)
else
  Result := Stream_WriteSmallUnicodeString_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteSmallUCS4String_LE(Stream: TStream; const Value: UCS4String; Advance: Boolean = True): TMemSize;
var
  TrueLen:  TStrSize;
begin
If Length(Value) > 0 then
  begin
    If Value[High(Value)] = 0 then
      TrueLen := Pred(Length(Value))
    else
      TrueLen := Length(Value);
    If TrueLen > High(Int16) then
      raise EBSStringTooLong.CreateFmt('Stream_WriteSmallUCS4String_LE: String is too long (%d) for small storage.',[Length(Value)])
    else If TrueLen > 0 then
      begin
        Result := Stream_WriteInt16_LE(Stream,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Stream_WriteUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Value[Low(Value)])),TrueLen));
      {$ELSE}
        Inc(Result,Stream_WriteBuffer_LE(Stream,Addr(Value[Low(Value)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ENDIF}
      end
    else Result := Stream_WriteInt16_LE(Stream,0,True);
  end
else Result := Stream_WriteInt16_LE(Stream,0,True);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUCS4String_BE(Stream: TStream; const Value: UCS4String; Advance: Boolean = True): TMemSize;
var
  TrueLen:  TStrSize;
begin
If Length(Value) > 0 then
  begin
    If Value[High(Value)] = 0 then
      TrueLen := Pred(Length(Value))
    else
      TrueLen := Length(Value);
    If TrueLen > High(Int16) then
      raise EBSStringTooLong.CreateFmt('Stream_WriteSmallUCS4String_BE: String is too long (%d) for small storage.',[Length(Value)])
    else If TrueLen > 0 then
      begin
        Result := Stream_WriteInt16_BE(Stream,TrueLen,True);
      {$IFDEF ENDIAN_BIG}
        Inc(Result,Stream_WriteBuffer_BE(Stream,Addr(Value[Low(Value)])^,TrueLen * SizeOf(UCS4Char),True));
      {$ELSE}
        Inc(Result,Stream_WriteUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Value[Low(Value)])),TrueLen));
      {$ENDIF}
      end
    else Result := Stream_WriteInt16_BE(Stream,0,True);
  end
else Result := Stream_WriteInt16_BE(Stream,0,True);
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallUCS4String(Stream: TStream; const Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallUCS4String_BE(Stream,Value,Advance)
else
  Result := Stream_WriteSmallUCS4String_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteSmallUCS4String(Stream: TStream; const Value: UCS4String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_WriteSmallUCS4String_BE(Stream,Value)
else
  Result := Stream_WriteSmallUCS4String_LE(Stream,Value);
end;

//==============================================================================

Function Stream_WriteSmallString_LE(Stream: TStream; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteSmallUTF8String_LE(Stream,StrToUTF8(Value),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallString_BE(Stream: TStream; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := Stream_WriteSmallUTF8String_BE(Stream,StrToUTF8(Value),Advance);
end;

//------------------------------------------------------------------------------

Function Stream_WriteSmallString(Stream: TStream; const Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteSmallUTF8String(Stream,StrToUTF8(Value),Advance,Endian);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_WriteSmallString(Stream: TStream; const Value: String; Endian: TEndian = endDefault): TMemSize;
begin
Result := Stream_WriteSmallUTF8String(Stream,StrToUTF8(Value),Endian);
end;

{<lite-begin>}
{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

{<lite-replace _Stream Stream>}
Function _Stream_WriteBuffer(Stream: TStream; const Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Stream.WriteBuffer(Buffer,Size);
Result := Size;
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Stream Stream>}
Function _Stream_WriteBytes(Stream: TStream; const Value: array of UInt8; Advance: Boolean): TMemSize;
var
  Buffer:   packed array[0..1023] of UInt8; // buffer is on stack, keep it small
  Remain:   Integer;
  Offset:   Integer;
  CopyCnt:  Integer;
  i:        Integer;
begin
If not OpenByteArrayIsPacked then
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
else If Length(Value) > 0 then
  Stream.WriteBuffer(Value[Low(Value)],Length(Value));
Result := TMemSize(Length(Value));
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Stream Stream>}
Function _Stream_FillBytes(Stream: TStream; Count: TMemSize; Value: UInt8; Advance: Boolean): TMemSize;

  Function Min(A,B: TMemSize): TMemSize;
  begin
    If A < B then
      Result := A
    else
      Result := B;
  end;

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
    CopyCnt := Min(Count,TMemSize(SizeOf(Buffer)));
    Stream.WriteBuffer(Buffer,CopyCnt);
    Dec(Count,TMemSize(CopyCnt));
    Inc(Result,TMemSize(CopyCnt));
  end;
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-lns 2>}

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

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function Stream_WriteVariant_LE(Stream: TStream; const Value: Variant; Advance: Boolean = True): TMemSize;
{$DEFINE BS_INC_VW}{$DEFINE BS_INC_L}
  {$INCLUDE '.\BinaryStreaming_var.inc'}
{$UNDEF BS_INC_VW}{$UNDEF BS_INC_L}
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function Stream_WriteVariant_BE(Stream: TStream; const Value: Variant; Advance: Boolean = True): TMemSize;
{$DEFINE BS_INC_VW}
  {$INCLUDE '.\BinaryStreaming_var.inc'}
{$UNDEF BS_INC_VW}
{$IFDEF FPCDWM}{$POP}{$ENDIF}

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


{<lite-begin>}
{===============================================================================
--------------------------------------------------------------------------------
                                 Stream reading
--------------------------------------------------------------------------------
===============================================================================}
{-------------------------------------------------------------------------------
    Booleans
-------------------------------------------------------------------------------}

{<lite-replace _Stream Stream>}
Function _Stream_ReadBool(Stream: TStream; out Value: ByteBool; Advance: Boolean): TMemSize;
var
  Temp: UInt8;
begin
Stream.ReadBuffer(Addr(Temp)^,SizeOf(Temp));
Result := SizeOf(Temp);
Value := NumToBool(Temp);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_GetBool_LE(Stream: TStream; Advance: Boolean = True): ByteBool;
begin
{<lite-replace _Stream Stream>}
_Stream_ReadBool(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadBoolean_LE(Stream: TStream; out Value: Boolean; Advance: Boolean = True): TMemSize;
var
  TempBool: ByteBool;
begin
Result := Stream_ReadBool_LE(Stream,TempBool,Advance);
Value := TempBool;
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetBoolean_LE(Stream: TStream; Advance: Boolean = True): Boolean;
begin
Result := Stream_GetBool_LE(Stream,Advance);
end;
{<lite-end-ln>}

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
        
{<lite-begin>}
{-------------------------------------------------------------------------------
    Integers
-------------------------------------------------------------------------------}

{<lite-replace _Stream Stream>}
Function _Stream_ReadInt8(Stream: TStream; out Value: Int8; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_GetInt8_LE(Stream: TStream; Advance: Boolean = True): Int8;
begin
{<lite-replace _Stream Stream>}
_Stream_ReadInt8(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Stream Stream>}
Function _Stream_ReadUInt8(Stream: TStream; out Value: UInt8; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_GetUInt8_LE(Stream: TStream; Advance: Boolean = True): UInt8;
begin
{<lite-replace _Stream Stream>}
_Stream_ReadUInt8(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadInt16_LE(Stream: TStream; out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := Int16(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetInt16_LE(Stream: TStream; Advance: Boolean = True): Int16;
begin
Stream_ReadInt16_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadUInt16_LE(Stream: TStream; out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetUInt16_LE(Stream: TStream; Advance: Boolean = True): UInt16;
begin
Stream_ReadUInt16_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadInt32_LE(Stream: TStream; out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := Int32(SwapEndian(UInt32(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetInt32_LE(Stream: TStream; Advance: Boolean = True): Int32;
begin
Stream_ReadInt32_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadUInt32_LE(Stream: TStream; out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetUInt32_LE(Stream: TStream; Advance: Boolean = True): UInt32;
begin
Stream_ReadUInt32_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadInt64_LE(Stream: TStream; out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := Int64(SwapEndian(UInt64(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetInt64_LE(Stream: TStream; Advance: Boolean = True): Int64;
begin
Stream_ReadInt64_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadUInt64_LE(Stream: TStream; out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := SwapEndian(Value);
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetUInt64_LE(Stream: TStream; Advance: Boolean = True): UInt64;
begin
Stream_ReadUInt64_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Floating point numbers
-------------------------------------------------------------------------------}

{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetFloat32_LE(Stream: TStream; Advance: Boolean = True): Float32;
begin
Stream_ReadFloat32_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetFloat64_LE(Stream: TStream; Advance: Boolean = True): Float64;
begin
Stream_ReadFloat64_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetFloat80_LE(Stream: TStream; Advance: Boolean = True): Float80;
begin
Stream_ReadFloat80_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadDateTime_LE(Stream: TStream; out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := Stream_ReadFloat64_LE(Stream,Float64(Value),Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetDateTime_LE(Stream: TStream; Advance: Boolean = True): TDateTime;
begin
Result := Stream_GetFloat64_LE(Stream,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
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
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetCurrency_LE(Stream: TStream; Advance: Boolean = True): Currency;
begin
Stream_ReadCurrency_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Characters
-------------------------------------------------------------------------------}

{<lite-replace _Stream Stream>}
Function _Stream_ReadAnsiChar(Stream: TStream; out Value: AnsiChar; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_GetAnsiChar_LE(Stream: TStream; Advance: Boolean = True): AnsiChar;
begin
{<lite-replace _Stream Stream>}
_Stream_ReadAnsiChar(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _Stream Stream>}
Function _Stream_ReadUTF8Char(Stream: TStream; out Value: UTF8Char; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_GetUTF8Char_LE(Stream: TStream; Advance: Boolean = True): UTF8Char;
begin
{<lite-replace _Stream Stream>}
_Stream_ReadUTF8Char(Stream,Result,Advance);
end;
{<lite-end-ln>}

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
 
{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadWideChar_LE(Stream: TStream; out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := WideChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetWideChar_LE(Stream: TStream; Advance: Boolean = True): WideChar;
begin
Stream_ReadWideChar_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadUnicodeChar_LE(Stream: TStream; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := UnicodeChar(SwapEndian(UInt16(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetUnicodeChar_LE(Stream: TStream; Advance: Boolean = True): UnicodeChar;
begin
Stream_ReadUnicodeChar_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-replace _LE "">}
Function Stream_ReadUCS4Char_LE(Stream: TStream; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Stream.ReadBuffer(Addr(Value)^,SizeOf(Value));
{$IFDEF ENDIAN_BIG}
Value := UCS4Char(SwapEndian(UInt32(Value)));
{$ENDIF}
Result := SizeOf(Value);
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetUCS4Char_LE(Stream: TStream; Advance: Boolean = True): UCS4Char;
begin
Stream_ReadUCS4Char_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>} 
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadChar_LE(Stream: TStream; out Value: Char; Advance: Boolean = True): TMemSize;
var
  Temp: UInt16;
begin
Result := Stream_ReadUInt16_LE(Stream,Temp,Advance);
Value := Char(Temp);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetChar_LE(Stream: TStream; Advance: Boolean = True): Char;
begin
Result := Char(Stream_GetUInt16_LE(Stream,Advance));
end;
{<lite-end-ln>}

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

{<lite-begin>}
{-------------------------------------------------------------------------------
    Strings
-------------------------------------------------------------------------------}

{<lite-block-replace-begin _LE "">}
{<lite-replace _Stream Stream>}
Function _Stream_ReadShortString(Stream: TStream; out Value: ShortString; Advance: Boolean): TMemSize;
var
  StrLength:  UInt8;
begin
Result := Stream_ReadUInt8_LE(Stream,StrLength,True);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_LE(Stream,Addr(Value[1])^,StrLength,True));
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_ReadShortString_LE(Stream: TStream; out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadShortString(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadShortString_BE(Stream: TStream; out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := _Stream_ReadShortString(Stream,Value,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_ReadShortString(Stream: TStream; out Value: ShortString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadShortString_BE(Stream,Value,Advance)
else
  Result := Stream_ReadShortString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadShortString(Stream: TStream; out Value: ShortString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadShortString_BE(Stream,Value)
else
  Result := Stream_ReadShortString_LE(Stream,Value);
end;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-replace _LE "">}
Function Stream_GetShortString_LE(Stream: TStream; Advance: Boolean = True): ShortString;
begin
{<lite-replace _Stream Stream>}
_Stream_ReadShortString(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadAnsiString_LE(Stream: TStream; out Value: AnsiString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_LE(Stream,PAnsiChar(Value)^,StrLength * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_ReadAnsiString_BE(Stream: TStream; out Value: AnsiString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_BE(Stream,PAnsiChar(Value)^,StrLength * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadAnsiString(Stream: TStream; out Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadAnsiString_BE(Stream,Value,Advance)
else
  Result := Stream_ReadAnsiString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadAnsiString(Stream: TStream; out Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadAnsiString_BE(Stream,Value)
else
  Result := Stream_ReadAnsiString_LE(Stream,Value);
end;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetAnsiString_LE(Stream: TStream; Advance: Boolean = True): AnsiString;
begin
Stream_ReadAnsiString_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadUTF8String_LE(Stream: TStream; out Value: UTF8String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_LE(Stream,PUTF8Char(Value)^,StrLength * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_ReadUTF8String_BE(Stream: TStream; out Value: UTF8String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_BE(Stream,PUTF8Char(Value)^,StrLength * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUTF8String(Stream: TStream; out Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUTF8String_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUTF8String_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUTF8String(Stream: TStream; out Value: UTF8String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUTF8String_BE(Stream,Value)
else
  Result := Stream_ReadUTF8String_LE(Stream,Value);
end;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetUTF8String_LE(Stream: TStream; Advance: Boolean = True): UTF8String;
begin
Stream_ReadUTF8String_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadWideString_LE(Stream: TStream; out Value: WideString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Value)),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,PWideChar(Value)^,StrLength * SizeOf(WideChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_ReadWideString_BE(Stream: TStream; out Value: WideString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,PWideChar(Value)^,StrLength * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Value)),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadWideString(Stream: TStream; out Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadWideString_BE(Stream,Value,Advance)
else
  Result := Stream_ReadWideString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadWideString(Stream: TStream; out Value: WideString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadWideString_BE(Stream,Value)
else
  Result := Stream_ReadWideString_LE(Stream,Value);
end;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetWideString_LE(Stream: TStream; Advance: Boolean = True): WideString;
begin
Stream_ReadWideString_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadUnicodeString_LE(Stream: TStream; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Value)),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,PUnicodeChar(Value)^,StrLength * SizeOf(UnicodeChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_ReadUnicodeString_BE(Stream: TStream; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,PUnicodeChar(Value)^,StrLength * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Value)),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUnicodeString(Stream: TStream; out Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUnicodeString_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUnicodeString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUnicodeString(Stream: TStream; out Value: UnicodeString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUnicodeString_BE(Stream,Value)
else
  Result := Stream_ReadUnicodeString_LE(Stream,Value);
end;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetUnicodeString_LE(Stream: TStream; Advance: Boolean = True): UnicodeString;
begin
Stream_ReadUnicodeString_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadUCS4String_LE(Stream: TStream; out Value: UCS4String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := nil;
SetLength(Value,StrLength + 1);
Value[High(Value)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Value[Low(Value)])),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,Addr(Value[Low(Value)])^,StrLength * SizeOf(UCS4Char),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_ReadUCS4String_BE(Stream: TStream; out Value: UCS4String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int32;
begin
Result := Stream_ReadInt32_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := nil;
SetLength(Value,StrLength + 1);
Value[High(Value)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,Addr(Value[Low(Value)])^,StrLength * SizeOf(UCS4Char),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Value[Low(Value)])),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadUCS4String(Stream: TStream; out Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUCS4String_BE(Stream,Value,Advance)
else
  Result := Stream_ReadUCS4String_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadUCS4String(Stream: TStream; out Value: UCS4String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadUCS4String_BE(Stream,Value)
else
  Result := Stream_ReadUCS4String_LE(Stream,Value);
end;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetUCS4String_LE(Stream: TStream; Advance: Boolean = True): UCS4String;
begin
Stream_ReadUCS4String_LE(Stream,Result,Advance);
end;
{<lite-end-ln>}

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

{<lite-begin>}
//==============================================================================

{<lite-block-replace-begin _LE "">}
Function Stream_ReadString_LE(Stream: TStream; out Value: String; Advance: Boolean = True): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadUTF8String_LE(Stream,TempStr,Advance);
Value := UTF8ToStr(TempStr);
end;
{<lite-end-ln>}

//------------------------------------------------------------------------------

Function Stream_ReadString_BE(Stream: TStream; out Value: String; Advance: Boolean = True): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadUTF8String_BE(Stream,TempStr,Advance);
Value := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Stream_ReadString(Stream: TStream; out Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadUTF8String(Stream,TempStr,Advance,Endian);
Value := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadString(Stream: TStream; out Value: String; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadUTF8String(Stream,TempStr,Endian);
Value := UTF8ToStr(TempStr);
end;

{<lite-begin>}
//------------------------------------------------------------------------------

{<lite-block-replace-begin _LE "">}
Function Stream_GetString_LE(Stream: TStream; Advance: Boolean = True): String;
begin
Result := UTF8ToStr(Stream_GetUTF8String_LE(Stream,Advance));
end;
{<lite-end-ln>}

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
    Small strings
-------------------------------------------------------------------------------}

Function Stream_ReadSmallAnsiString_LE(Stream: TStream; out Value: AnsiString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_LE(Stream,PAnsiChar(Value)^,StrLength * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallAnsiString_BE(Stream: TStream; out Value: AnsiString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_BE(Stream,PAnsiChar(Value)^,StrLength * SizeOf(AnsiChar),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallAnsiString(Stream: TStream; out Value: AnsiString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallAnsiString_BE(Stream,Value,Advance)
else
  Result := Stream_ReadSmallAnsiString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadSmallAnsiString(Stream: TStream; out Value: AnsiString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallAnsiString_BE(Stream,Value)
else
  Result := Stream_ReadSmallAnsiString_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallAnsiString_LE(Stream: TStream; Advance: Boolean = True): AnsiString;
begin
Stream_ReadSmallAnsiString_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallAnsiString_BE(Stream: TStream; Advance: Boolean = True): AnsiString;
begin
Stream_ReadSmallAnsiString_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallAnsiString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): AnsiString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallAnsiString_BE(Stream,Advance)
else
  Result := Stream_GetSmallAnsiString_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetSmallAnsiString(Stream: TStream; Endian: TEndian = endDefault): AnsiString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallAnsiString_BE(Stream)
else
  Result := Stream_GetSmallAnsiString_LE(Stream);
end;

//==============================================================================

Function Stream_ReadSmallUTF8String_LE(Stream: TStream; out Value: UTF8String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_LE(Stream,PUTF8Char(Value)^,StrLength * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;


//------------------------------------------------------------------------------

Function Stream_ReadSmallUTF8String_BE(Stream: TStream; out Value: UTF8String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
  Inc(Result,Stream_ReadBuffer_BE(Stream,PUTF8Char(Value)^,StrLength * SizeOf(UTF8Char),True));
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallUTF8String(Stream: TStream; out Value: UTF8String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallUTF8String_BE(Stream,Value,Advance)
else
  Result := Stream_ReadSmallUTF8String_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadSmallUTF8String(Stream: TStream; out Value: UTF8String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallUTF8String_BE(Stream,Value)
else
  Result := Stream_ReadSmallUTF8String_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUTF8String_LE(Stream: TStream; Advance: Boolean = True): UTF8String;
begin
Stream_ReadSmallUTF8String_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUTF8String_BE(Stream: TStream; Advance: Boolean = True): UTF8String;
begin
Stream_ReadSmallUTF8String_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUTF8String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UTF8String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallUTF8String_BE(Stream,Advance)
else
  Result := Stream_GetSmallUTF8String_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetSmallUTF8String(Stream: TStream; Endian: TEndian = endDefault): UTF8String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallUTF8String_BE(Stream)
else
  Result := Stream_GetSmallUTF8String_LE(Stream);
end;

//==============================================================================

Function Stream_ReadSmallWideString_LE(Stream: TStream; out Value: WideString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Value)),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,PWideChar(Value)^,StrLength * SizeOf(WideChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallWideString_BE(Stream: TStream; out Value: WideString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,PWideChar(Value)^,StrLength * SizeOf(WideChar),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PWideChar(Value)),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallWideString(Stream: TStream; out Value: WideString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallWideString_BE(Stream,Value,Advance)
else
  Result := Stream_ReadSmallWideString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadSmallWideString(Stream: TStream; out Value: WideString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallWideString_BE(Stream,Value)
else
  Result := Stream_ReadSmallWideString_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallWideString_LE(Stream: TStream; Advance: Boolean = True): WideString;
begin
Stream_ReadSmallWideString_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallWideString_BE(Stream: TStream; Advance: Boolean = True): WideString;
begin
Stream_ReadSmallWideString_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallWideString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): WideString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallWideString_BE(Stream,Advance)
else
  Result := Stream_GetSmallWideString_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetSmallWideString(Stream: TStream; Endian: TEndian = endDefault): WideString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallWideString_BE(Stream)
else
  Result := Stream_GetSmallWideString_LE(Stream);
end;

//==============================================================================

Function Stream_ReadSmallUnicodeString_LE(Stream: TStream; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Value)),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,PUnicodeChar(Value)^,StrLength * SizeOf(UnicodeChar),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallUnicodeString_BE(Stream: TStream; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := '';
SetLength(Value,StrLength);
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,PUnicodeChar(Value)^,StrLength * SizeOf(UnicodeChar),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt16Arr_SwapEndian(Stream,PUInt16(PUnicodeChar(Value)),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallUnicodeString(Stream: TStream; out Value: UnicodeString; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallUnicodeString_BE(Stream,Value,Advance)
else
  Result := Stream_ReadSmallUnicodeString_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadSmallUnicodeString(Stream: TStream; out Value: UnicodeString; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallUnicodeString_BE(Stream,Value)
else
  Result := Stream_ReadSmallUnicodeString_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUnicodeString_LE(Stream: TStream; Advance: Boolean = True): UnicodeString;
begin
Stream_ReadSmallUnicodeString_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUnicodeString_BE(Stream: TStream; Advance: Boolean = True): UnicodeString;
begin
Stream_ReadSmallUnicodeString_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUnicodeString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UnicodeString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallUnicodeString_BE(Stream,Advance)
else
  Result := Stream_GetSmallUnicodeString_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetSmallUnicodeString(Stream: TStream; Endian: TEndian = endDefault): UnicodeString;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallUnicodeString_BE(Stream)
else
  Result := Stream_GetSmallUnicodeString_LE(Stream);
end;

//==============================================================================

Function Stream_ReadSmallUCS4String_LE(Stream: TStream; out Value: UCS4String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_LE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := nil;
SetLength(Value,StrLength + 1);
Value[High(Value)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Value[Low(Value)])),StrLength));
{$ELSE}
  Inc(Result,Stream_ReadBuffer_LE(Stream,Addr(Value[Low(Value)])^,StrLength * SizeOf(UCS4Char),True));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallUCS4String_BE(Stream: TStream; out Value: UCS4String; Advance: Boolean = True): TMemSize;
var
  StrLength:  Int16;
begin
Result := Stream_ReadInt16_BE(Stream,StrLength,True);
ClampStringLength(StrLength);
Value := nil;
SetLength(Value,StrLength + 1);
Value[High(Value)] := 0;
If StrLength > 0 then
{$IFDEF ENDIAN_BIG}
  Inc(Result,Stream_ReadBuffer_BE(Stream,Addr(Value[Low(Value)])^,StrLength * SizeOf(UCS4Char),True));
{$ELSE}
  Inc(Result,Stream_ReadUInt32Arr_SwapEndian(Stream,PUInt32(Addr(Value[Low(Value)])),StrLength));
{$ENDIF}
AdvanceStream(Advance,Stream,Result);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallUCS4String(Stream: TStream; out Value: UCS4String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallUCS4String_BE(Stream,Value,Advance)
else
  Result := Stream_ReadSmallUCS4String_LE(Stream,Value,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadSmallUCS4String(Stream: TStream; out Value: UCS4String; Endian: TEndian = endDefault): TMemSize;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_ReadSmallUCS4String_BE(Stream,Value)
else
  Result := Stream_ReadSmallUCS4String_LE(Stream,Value);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUCS4String_LE(Stream: TStream; Advance: Boolean = True): UCS4String;
begin
Stream_ReadSmallUCS4String_LE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUCS4String_BE(Stream: TStream; Advance: Boolean = True): UCS4String;
begin
Stream_ReadSmallUCS4String_BE(Stream,Result,Advance);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallUCS4String(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): UCS4String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallUCS4String_BE(Stream,Advance)
else
  Result := Stream_GetSmallUCS4String_LE(Stream,Advance);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetSmallUCS4String(Stream: TStream; Endian: TEndian = endDefault): UCS4String;
begin
If ResolveEndian(Endian) = endBig then
  Result := Stream_GetSmallUCS4String_BE(Stream)
else
  Result := Stream_GetSmallUCS4String_LE(Stream);
end;

//==============================================================================

Function Stream_ReadSmallString_LE(Stream: TStream; out Value: String; Advance: Boolean = True): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadSmallUTF8String_LE(Stream,TempStr,Advance);
Value := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallString_BE(Stream: TStream; out Value: String; Advance: Boolean = True): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadSmallUTF8String_BE(Stream,TempStr,Advance);
Value := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Stream_ReadSmallString(Stream: TStream; out Value: String; Advance: Boolean; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadSmallUTF8String(Stream,TempStr,Advance,Endian);
Value := UTF8ToStr(TempStr);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_ReadSmallString(Stream: TStream; out Value: String; Endian: TEndian = endDefault): TMemSize;
var
  TempStr:  UTF8String;
begin
Result := Stream_ReadSmallUTF8String(Stream,TempStr,Endian);
Value := UTF8ToStr(TempStr);
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallString_LE(Stream: TStream; Advance: Boolean = True): String;
begin
Result := UTF8ToStr(Stream_GetSmallUTF8String_LE(Stream,Advance));
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallString_BE(Stream: TStream; Advance: Boolean = True): String;
begin
Result := UTF8ToStr(Stream_GetSmallUTF8String_BE(Stream,Advance));
end;

//------------------------------------------------------------------------------

Function Stream_GetSmallString(Stream: TStream; Advance: Boolean; Endian: TEndian = endDefault): String;
begin
Result := UTF8ToStr(Stream_GetSmallUTF8String(Stream,Advance,Endian));
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function Stream_GetSmallString(Stream: TStream; Endian: TEndian = endDefault): String;
begin
Result := UTF8ToStr(Stream_GetSmallUTF8String(Stream,Endian));
end;

{<lite-begin>}
{-------------------------------------------------------------------------------
    General data buffers
-------------------------------------------------------------------------------}

{<lite-replace _Stream Stream>}
Function _Stream_ReadBuffer(Stream: TStream; out Buffer; Size: TMemSize; Advance: Boolean): TMemSize;
begin
Stream.ReadBuffer(Addr(Buffer)^,Size);
Result := Size;
AdvanceStream(Advance,Stream,Result);
end;
{<lite-end-lns 2>}

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

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function Stream_ReadVariant_LE(Stream: TStream; out Value: Variant; Advance: Boolean = True): TMemSize;
{$DEFINE BS_INC_VR}{$DEFINE BS_INC_L}
  {$INCLUDE '.\BinaryStreaming_var.inc'}
{$UNDEF BS_INC_VR}{$UNDEF BS_INC_L}
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W6058{$ENDIF}
Function Stream_ReadVariant_BE(Stream: TStream; out Value: Variant; Advance: Boolean = True): TMemSize;
{$DEFINE BS_INC_VR}
  {$INCLUDE '.\BinaryStreaming_var.inc'}
{$UNDEF BS_INC_VR}
{$IFDEF FPCDWM}{$POP}{$ENDIF}

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
const
  BS_VALTYPE_BOOL           = 0;
  BS_VALTYPE_BOOLEAN        = 1;
  BS_VALTYPE_NUM_1          = 2;
  BS_VALTYPE_NUM_2          = 3;
  BS_VALTYPE_NUM_4          = 4;
  BS_VALTYPE_NUM_8          = 5;
  BS_VALTYPE_NUM_10         = 6;
  BS_VALTYPE_CHAR           = 7;
  BS_VALTYPE_STR_SHORT      = 8;
  BS_VALTYPE_STR_ANSI       = 9;
  BS_VALTYPE_STR_UTF8       = 10;
  BS_VALTYPE_STR_WIDE       = 11;
  BS_VALTYPE_STR_UNICODE    = 12;
  BS_VALTYPE_STR_UCS4       = 13;
  BS_VALTYPE_STR            = 14;
  BS_VALTYPE_SMLSTR_ANSI    = 15;
  BS_VALTYPE_SMLSTR_UTF8    = 16;
  BS_VALTYPE_SMLSTR_WIDE    = 17;
  BS_VALTYPE_SMLSTR_UNICODE = 18;
  BS_VALTYPE_SMLSTR_UCS4    = 19;
  BS_VALTYPE_SMLSTR         = 20;
  
  BS_VALTYPE_BUFFER         = 27;
  BS_VALTYPE_BYTEARRAY      = 28;
  BS_VALTYPE_FILL           = 29;
  BS_VALTYPE_VARIANT        = 30;

//------------------------------------------------------------------------------
{===============================================================================
    TCustomStreamer - class implementation
===============================================================================}
{-------------------------------------------------------------------------------
    TCustomStreamer - protected methods
-------------------------------------------------------------------------------}

{$IFOPT Q+}{$DEFINE BS_OverflowChecks}{$Q-}{$ENDIF}

Function TCustomStreamer.GetOffset: Int64;
begin
Result := Position - fStart;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.SetOffset(NewOffset: Int64);
begin
Position := fStart + NewOffset;
end;

{$IFDEF BS_OverflowChecks}{$UNDEF BS_OverflowChecks}{$Q+}{$ENDIF}

//------------------------------------------------------------------------------

Function TCustomStreamer.GetBookmark(Index: Integer): TBSBookmarkData;
begin
If BookmarkCheckIndex(Index) then
  Result := fBookmarks[Index]
else
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.GetBookmark: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.SetBookmark(Index: Integer; Value: TBSBookmarkData);
begin
If BookmarkCheckIndex(Index) then
  fBookmarks[Index] := Value
else
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.SetBookmark: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetBookmarkPtr(Index: Integer): PBSBookmarkData;
begin
If BookmarkCheckIndex(Index) then
  Result := Addr(fBookmarks[Index])
else
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.GetBookmarkPtr: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCapacity: Integer;
begin
Result := Length(fBookmarks);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.SetCapacity(Value: Integer);
begin
If Value >= 0 then
  begin
    If Value <> Length(fBookmarks) then
      begin
        SetLength(fBookMarks,Value);
        If Value < fBookmarkCount then
          fBookmarkCount := Value;
      end;
  end
else raise EBSInvalidValue.CreateFmt('TCustomStreamer.SetCapacity: Invalid capacity (%d).',[Value]);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCount: Integer;
begin
Result := fBookmarkCount;
end;

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W5024{$ENDIF}
procedure TCustomStreamer.SetCount(Value: Integer);
begin
// do nothing
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteOpenByteArray(Ptr: Pointer; Count: TMemSize; Advance: Boolean): TMemSize;
var
  OldPos:   Int64;
  WorkPtr:  PUInt8;
  i:        Integer;
begin
Result := 0;
If not OpenByteArrayIsPacked then
  begin
    OldPos := Position;
    WorkPtr := Ptr;
    For i := 0 to Pred(Count) do
      begin
        Inc(Result,WriteUInt8(UInt8(WorkPtr^),True));
        Inc(WorkPtr,OpenByteArrayStride);
      end;
    If not Advance then
      Position := OldPos;
  end
else If Count > 0 then
  Result := WriteValue(BS_VALTYPE_BUFFER,Ptr,Advance,Count);
{$message 'todo - rewisit and optimize (buffering)'}
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteValueAt(Position: Int64; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveAt(Position);
Result := WriteValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteValueAtOffset(Offset: Int64; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveAtOffset(Offset);
Result := WriteValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteValueTo(ID: TBSBookmarkID; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveTo(ID);
Result := WriteValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteValueToIndex(Index: Integer; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveToIndex(Index);
Result := WriteValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadValueAt(Position: Int64; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveAt(Position);
Result := ReadValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadValueAtOffset(Offset: Int64; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveAtOffset(Offset);
Result := ReadValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadValueFrom(ID: TBSBookmarkID; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveTo(ID);
Result := ReadValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadValueFromIndex(Index: Integer; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveToIndex(Index);
Result := ReadValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.DoChange;
begin
If fChangingCounter <= 0 then
  begin
    If Assigned(fOnChangeEvent) then
      fOnChangeEvent(Self)
    else If Assigned(fOnChangeCallback) then
      fOnChangeCallback(Self);
    fChanged := False;  
  end
else fChanged := True;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.Initialize;
begin
fEndian := endDefault;
fStart := 0;
SetLength(fBookmarks,0);
fBookmarkCount := 0; 
fChanged := False;
fChangingCounter := 0;
fOnChangeEvent := nil;
fOnChangeCallback := nil;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.Finalize;
begin
// prevent events firing
fOnChangeEvent := nil;
fOnChangeCallback := nil;
BookmarkClear;
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

Function TCustomStreamer.BeginUpdate: Integer;
begin
Inc(fChangingCounter);
Result := fChangingCounter;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.EndUpdate: Integer;
begin
Dec(fChangingCounter);
If fChangingCounter <= 0 then
  begin
    fChangingCounter := 0;
    If fChanged then
      DoChange; // sets fChanged to false
  end;
Result := fChangingCounter;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.LowIndex: Integer;
begin
Result := Low(fBookmarks);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.HighIndex: Integer;
begin
Result := Pred(fBookmarkCount);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkLowIndex: Integer;
begin
Result := LowIndex;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkHighIndex: Integer;
begin
Result := highIndex;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkCheckIndex(Index: Integer): Boolean;
begin
Result := CheckIndex(Index);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkIndexOf(ID: TBSBookmarkID): Integer;
var
  i:  Integer;
begin
Result := -1;
For i := BookmarkLowIndex to BookmarkHighIndex do
  If fBookmarks[i].ID = ID then
    begin
      Result := i;
      Break{For i};
    end;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkFind(ID: TBSBookmarkID; out Index: Integer): Boolean;
begin
Index := BookmarkIndexOf(ID);
Result := BookmarkCheckIndex(Index);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkAdd(ID: TBSBookmarkID; Position: Int64): Integer;
begin
If not BookmarkFind(ID,Result) then
  begin
    Grow;
    Result := fBookmarkCount;
    fBookmarks[Result].ID := ID;
    fBookmarks[Result].Position := Position;
    Inc(fBookmarkCount);
    DoChange;
  end
else raise EBSDuplicateItem.CreateFmt('TCustomStreamer.BookmarkAdd: Bookmark %d already exists.',[ID]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function TCustomStreamer.BookmarkAdd(ID: TBSBookmarkID): Integer;
begin
Result := BookmarkAdd(ID,Position);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.BookmarkInsert(Index: Integer; ID: TBSBookmarkID; Position: Int64);
var
  i:  Integer;
begin
If not BookmarkFind(ID,i) then
  begin
    If CheckIndex(Index) then
      begin
        Grow;
        For i := BookmarkHighIndex downto Index do
          fBookmarks[i + 1] := fBookmarks[i];
        fBookmarks[Index].ID := ID;
        fBookmarks[Index].Position := Position;
        Inc(fBookmarkCount);
        DoChange;
      end
    else If Index = BookmarkCount then
      BookmarkAdd(ID,Position)
    else
      raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.BookmarkInsert: Index (%d) out of bounds.',[Index]);
  end
else raise EBSDuplicateItem.CreateFmt('TCustomStreamer.BookmarkInsert: Bookmark %d already exists.',[ID]);
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

procedure TCustomStreamer.BookmarkInsert(Index: Integer; ID: TBSBookmarkID);
begin
BookmarkInsert(Index,ID,Position);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.BookmarkMove(SrcIndex,DstIndex: Integer);
var
  Temp: TBSBookmarkData;
  i:    Integer;
begin
If not CheckIndex(SrcIndex) then
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.BookmarkInsert: Source index (%d) out of bounds.',[SrcIndex]);
If not CheckIndex(DstIndex) then
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.BookmarkInsert: Destination index (%d) out of bounds.',[DstIndex]);
If SrcIndex <> DstIndex then
  begin
    Temp := fBookmarks[SrcIndex];
    If SrcIndex < DstIndex then
      For i := SrcIndex to Pred(DstIndex) do
        fBookmarks[i] := fBookmarks[i + 1]
    else
      For i := SrcIndex downto Succ(DstIndex) do
        fBookmarks[i] := fBookmarks[i - 1];
    fBookmarks[DstIndex] := Temp;
    DoChange;
  end;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.BookmarkExchange(Index1,Index2: Integer);
var
  Temp: TBSBookmarkData;
begin
If not CheckIndex(Index1) then
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.BookmarkInsert: Index 1 (%d) out of bounds.',[Index1]);
If not CheckIndex(Index2) then
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.BookmarkInsert: Index 2 (%d) out of bounds.',[Index2]);
If Index1 <> Index2 then
  begin
    Temp := fBookmarks[Index1];
    fBookmarks[Index1] := fBookmarks[Index2];
    fBookmarks[Index2] := Temp;
    DoChange;
  end;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkExtract(ID: TBSBookmarkID): TBSBookmarkData;
var
  Index:  Integer;
begin
If BookmarkFind(ID,Index) then
  begin
    Result := fBookmarks[Index];
    BookmarkDelete(Index);
  end
else raise EBSUnknownItem.CreateFmt('TCustomStreamer.BookmarkExtract: Item %d not found.',[ID]);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkRemove(ID: TBSBookmarkID): Integer;
begin
If BookmarkFind(ID,Result) then
  BookmarkDelete(Result)
else
  Result := -1;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.BookmarkDelete(Index: Integer);
var
  i:  Integer;
begin
If BookmarkCheckIndex(Index) then
  begin
    For i := Index to Pred(BookmarkHighIndex) do
      fBookmarks[i] := fBookmarks[i + 1];
    Dec(fBookmarkCount);
    Shrink;
    DoChange;
  end
else raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.BookmarkDelete: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.BookmarkClear;
begin
SetLength(fBookmarks,0);
fBookmarkCount := 0;
DoChange;
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.BookmarkGetPosition(ID: TBSBookmarkID): Int64;
var
  Index:  Integer;
begin
If BookmarkFind(ID,Index) then
  Result := fBookmarks[Index].Position
else
  raise EBSUnknownItem.CreateFmt('TCustomStreamer.BookmarkGetPosition: Item %d not found.',[ID]);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.BookmarkSetPosition(ID: TBSBookmarkID; NewPosition: Int64);
var
  Index:  Integer;
begin
If BookmarkFind(ID,Index) then
  fBookmarks[Index].Position := NewPosition
else
  raise EBSUnknownItem.CreateFmt('TCustomStreamer.BookmarkSetPosition: Item %d not found.',[ID]);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.MoveAt(Position: Int64);
begin
Self.Position := Position;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.MoveAtOffset(Offset: Int64);
begin
Self.Offset := Offset;
end;

//------------------------------------------------------------------------------

{$IFOPT Q+}{$DEFINE BS_OverflowChecks}{$Q-}{$ENDIF}
procedure TCustomStreamer.MoveBy(Delta: Int64);
begin
Position := Position + Delta;
end;
{$IFDEF BS_OverflowChecks}{$UNDEF BS_OverflowChecks}{$Q+}{$ENDIF}

//------------------------------------------------------------------------------

procedure TCustomStreamer.MoveToStart;
begin
Position := fStart;
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.MoveTo(ID: TBSBookmarkID);
var
  Index:  Integer;
begin
If BookmarkFind(ID,Index) then
  Position := fBookmarks[Index].Position
else
  raise EBSUnknownItem.CreateFmt('TCustomStreamer.MoveTo: Item %d not found.',[ID]);
end;

//------------------------------------------------------------------------------

procedure TCustomStreamer.MoveToIndex(Index: Integer);
begin
If BookmarkCheckIndex(Index) then
  Position := fBookmarks[Index].Position
else
  raise EBSIndexOutOfBounds.CreateFmt('TCustomStreamer.MoveToIndex: Index (%d) out of bounds.',[Index]);
end;

//==============================================================================

Function TCustomStreamer.WriteBool(Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBoolean(Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt8(Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt8(Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_1,@Value,Advance);
end;
 
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt16(Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt16(Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt32(Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt32(Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt64(Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt64(Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat32(Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat64(Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat80(Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteDateTime(Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCurrency(Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;   

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiChar(Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8Char(Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideChar(Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeChar(Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4Char(Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteChar(Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteShortString(const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiString(const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8String(const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideString(const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_STR_WIDE,@Value,Advance);
end;
 
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeString(const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4String(const Value: UCS4String; Advance: Boolean = True): TMemSize; 
begin
Result := WriteValue(BS_VALTYPE_STR_UCS4,@Value,Advance);
end;
  
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteString(const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallAnsiString(const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUTF8String(const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallWideString(const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;
 
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUnicodeString(const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUCS4String(const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;
  
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallString(const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_SMLSTR,@Value,Advance);
end;
  
//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBuffer(const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBytes(const Value: array of UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_BYTEARRAY,@Value,Advance,TMemSize(Length(Value)));
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.FillBytes(Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_FILL,@Value,Advance,Count);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteVariant(const Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := WriteValue(BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.WriteBoolAt(Position: Int64; Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBooleanAt(Position: Int64; Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt8At(Position: Int64; Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt8At(Position: Int64; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt16At(Position: Int64; Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt16At(Position: Int64; Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt32At(Position: Int64; Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt32At(Position: Int64; Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt64At(Position: Int64; Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt64At(Position: Int64; Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat32At(Position: Int64; Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat64At(Position: Int64; Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat80At(Position: Int64; Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteDateTimeAt(Position: Int64; Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCurrencyAt(Position: Int64; Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiCharAt(Position: Int64; Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8CharAt(Position: Int64; Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideCharAt(Position: Int64; Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeCharAt(Position: Int64; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4CharAt(Position: Int64; Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCharAt(Position: Int64; Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteShortStringAt(Position: Int64; const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiStringAt(Position: Int64; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8StringAt(Position: Int64; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideStringAt(Position: Int64; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeStringAt(Position: Int64; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4StringAt(Position: Int64; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteStringAt(Position: Int64; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallAnsiStringAt(Position: Int64; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUTF8StringAt(Position: Int64; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallWideStringAt(Position: Int64; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUnicodeStringAt(Position: Int64; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUCS4StringAt(Position: Int64; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallStringAt(Position: Int64; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBufferAt(Position: Int64; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBytesAt(Position: Int64; const Value: array of UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_BYTEARRAY,@Value,Advance,TMemSize(Length(Value)));
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.FillBytesAt(Position: Int64; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_FILL,@Value,Advance,Count);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteVariantAt(Position: Int64; const Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAt(Position,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.WriteBoolAtOffset(Offset: Int64; Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBooleanAtOffset(Offset: Int64; Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt8AtOffset(Offset: Int64; Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt8AtOffset(Offset: Int64; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt16AtOffset(Offset: Int64; Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt16AtOffset(Offset: Int64; Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt32AtOffset(Offset: Int64; Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt32AtOffset(Offset: Int64; Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt64AtOffset(Offset: Int64; Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt64AtOffset(Offset: Int64; Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat32AtOffset(Offset: Int64; Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat64AtOffset(Offset: Int64; Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat80AtOffset(Offset: Int64; Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteDateTimeAtOffset(Offset: Int64; Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCurrencyAtOffset(Offset: Int64; Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiCharAtOffset(Offset: Int64; Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8CharAtOffset(Offset: Int64; Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideCharAtOffset(Offset: Int64; Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeCharAtOffset(Offset: Int64; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4CharAtOffset(Offset: Int64; Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCharAtOffset(Offset: Int64; Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteShortStringAtOffset(Offset: Int64; const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiStringAtOffset(Offset: Int64; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8StringAtOffset(Offset: Int64; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideStringAtOffset(Offset: Int64; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeStringAtOffset(Offset: Int64; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4StringAtOffset(Offset: Int64; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteStringAtOffset(Offset: Int64; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallAnsiStringAtOffset(Offset: Int64; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUTF8StringAtOffset(Offset: Int64; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallWideStringAtOffset(Offset: Int64; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUnicodeStringAtOffset(Offset: Int64; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUCS4StringAtOffset(Offset: Int64; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallStringAtOffset(Offset: Int64; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBufferAtOffset(Offset: Int64; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBytesAtOffset(Offset: Int64; const Value: array of UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Position,BS_VALTYPE_BYTEARRAY,@Value,Advance,TMemSize(Length(Value)));
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.FillBytesAtOffset(Offset: Int64; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_FILL,@Value,Advance,Count);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteVariantAtOffset(Offset: Int64; const Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Offset,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.WriteBoolTo(ID: TBSBookmarkID; Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBooleanTo(ID: TBSBookmarkID; Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt8To(ID: TBSBookmarkID; Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt8To(ID: TBSBookmarkID; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt16To(ID: TBSBookmarkID; Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt16To(ID: TBSBookmarkID; Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt32To(ID: TBSBookmarkID; Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt32To(ID: TBSBookmarkID; Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt64To(ID: TBSBookmarkID; Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt64To(ID: TBSBookmarkID; Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat32To(ID: TBSBookmarkID; Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat64To(ID: TBSBookmarkID; Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat80To(ID: TBSBookmarkID; Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteDateTimeTo(ID: TBSBookmarkID; Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCurrencyTo(ID: TBSBookmarkID; Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiCharTo(ID: TBSBookmarkID; Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8CharTo(ID: TBSBookmarkID; Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideCharTo(ID: TBSBookmarkID; Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeCharTo(ID: TBSBookmarkID; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4CharTo(ID: TBSBookmarkID; Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCharTo(ID: TBSBookmarkID; Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteShortStringTo(ID: TBSBookmarkID; const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiStringTo(ID: TBSBookmarkID; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8StringTo(ID: TBSBookmarkID; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideStringTo(ID: TBSBookmarkID; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeStringTo(ID: TBSBookmarkID; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4StringTo(ID: TBSBookmarkID; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteStringTo(ID: TBSBookmarkID; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallAnsiStringTo(ID: TBSBookmarkID; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUTF8StringTo(ID: TBSBookmarkID; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallWideStringTo(ID: TBSBookmarkID; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUnicodeStringTo(ID: TBSBookmarkID; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUCS4StringTo(ID: TBSBookmarkID; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallStringTo(ID: TBSBookmarkID; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBufferTo(ID: TBSBookmarkID; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBytesTo(ID: TBSBookmarkID; const Value: array of UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Position,BS_VALTYPE_BYTEARRAY,@Value,Advance,TMemSize(Length(Value)));
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.FillBytesTo(ID: TBSBookmarkID; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_FILL,@Value,Advance,Count);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteVariantTo(ID: TBSBookmarkID; const Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueTo(ID,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.WriteBoolToIndex(Index: Integer; Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBooleanToIndex(Index: Integer; Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt8ToIndex(Index: Integer; Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt8ToIndex(Index: Integer; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt16ToIndex(Index: Integer; Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt16ToIndex(Index: Integer; Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt32ToIndex(Index: Integer; Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt32ToIndex(Index: Integer; Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteInt64ToIndex(Index: Integer; Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUInt64ToIndex(Index: Integer; Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat32ToIndex(Index: Integer; Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat64ToIndex(Index: Integer; Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteFloat80ToIndex(Index: Integer; Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteDateTimeToIndex(Index: Integer; Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCurrencyToIndex(Index: Integer; Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiCharToIndex(Index: Integer; Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8CharToIndex(Index: Integer; Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideCharToIndex(Index: Integer; Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeCharToIndex(Index: Integer; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4CharToIndex(Index: Integer; Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteCharToIndex(Index: Integer; Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteShortStringToIndex(Index: Integer; const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteAnsiStringToIndex(Index: Integer; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUTF8StringToIndex(Index: Integer; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteWideStringToIndex(Index: Integer; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUnicodeStringToIndex(Index: Integer; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteUCS4StringToIndex(Index: Integer; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteStringToIndex(Index: Integer; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallAnsiStringToIndex(Index: Integer; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUTF8StringToIndex(Index: Integer; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallWideStringToIndex(Index: Integer; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUnicodeStringToIndex(Index: Integer; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallUCS4StringToIndex(Index: Integer; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteSmallStringToIndex(Index: Integer; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBufferToIndex(Index: Integer; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteBytesToIndex(Index: Integer; const Value: array of UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtOffset(Position,BS_VALTYPE_BYTEARRAY,@Value,Advance,TMemSize(Length(Value)));
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.FillBytesToIndex(Index: Integer; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_FILL,@Value,Advance,Count);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.WriteVariantToIndex(Index: Integer; const Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueToIndex(Index,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.ReadBool(out Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBoolean(out Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt8(out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt8(out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt16(out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt16(out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt32(out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt32(out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt64(out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt64(out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat32(out Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat64(out Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat80(out Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadDateTime(out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCurrency(out Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiChar(out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8Char(out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideChar(out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeChar(out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4Char(out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadChar(out Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadShortString(out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiString(out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8String(out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideString(out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeString(out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4String(out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadString(out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallAnsiString(out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUTF8String(out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallWideString(out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUnicodeString(out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUCS4String(out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallString(out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBuffer(out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadVariant(out Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := ReadValue(BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.ReadBoolAt(Position: Int64; out Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBooleanAt(Position: Int64; out Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt8At(Position: Int64; out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt8At(Position: Int64; out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt16At(Position: Int64; out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt16At(Position: Int64; out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt32At(Position: Int64; out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt32At(Position: Int64; out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt64At(Position: Int64; out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt64At(Position: Int64; out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat32At(Position: Int64; out Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat64At(Position: Int64; out Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat80At(Position: Int64; out Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadDateTimeAt(Position: Int64; out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCurrencyAt(Position: Int64; out Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiCharAt(Position: Int64; out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8CharAt(Position: Int64; out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideCharAt(Position: Int64; out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeCharAt(Position: Int64; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4CharAt(Position: Int64; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCharAt(Position: Int64; out Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadShortStringAt(Position: Int64; out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiStringAt(Position: Int64; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8StringAt(Position: Int64; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideStringAt(Position: Int64; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeStringAt(Position: Int64; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4StringAt(Position: Int64; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadStringAt(Position: Int64; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallAnsiStringAt(Position: Int64; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUTF8StringAt(Position: Int64; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallWideStringAt(Position: Int64; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUnicodeStringAt(Position: Int64; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUCS4StringAt(Position: Int64; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallStringAt(Position: Int64; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBufferAt(Position: Int64; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadVariantAt(Position: Int64; out Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAt(Position,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.ReadBoolAtOffset(Offset: Int64; out Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBooleanAtOffset(Offset: Int64; out Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt8AtOffset(Offset: Int64; out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt8AtOffset(Offset: Int64; out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt16AtOffset(Offset: Int64; out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt16AtOffset(Offset: Int64; out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt32AtOffset(Offset: Int64; out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt32AtOffset(Offset: Int64; out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt64AtOffset(Offset: Int64; out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt64AtOffset(Offset: Int64; out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat32AtOffset(Offset: Int64; out Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat64AtOffset(Offset: Int64; out Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat80AtOffset(Offset: Int64; out Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadDateTimeAtOffset(Offset: Int64; out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCurrencyAtOffset(Offset: Int64; out Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiCharAtOffset(Offset: Int64; out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8CharAtOffset(Offset: Int64; out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideCharAtOffset(Offset: Int64; out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeCharAtOffset(Offset: Int64; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4CharAtOffset(Offset: Int64; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCharAtOffset(Offset: Int64; out Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadShortStringAtOffset(Offset: Int64; out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiStringAtOffset(Offset: Int64; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8StringAtOffset(Offset: Int64; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideStringAtOffset(Offset: Int64; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeStringAtOffset(Offset: Int64; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4StringAtOffset(Offset: Int64; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadStringAtOffset(Offset: Int64; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallAnsiStringAtOffset(Offset: Int64; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUTF8StringAtOffset(Offset: Int64; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallWideStringAtOffset(Offset: Int64; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUnicodeStringAtOffset(Offset: Int64; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUCS4StringAtOffset(Offset: Int64; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallStringAtOffset(Offset: Int64; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBufferAtOffset(Offset: Int64; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadVariantAtOffset(Offset: Int64; out Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtOffset(Offset,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.ReadBoolFrom(ID: TBSBookmarkID; out Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBooleanFrom(ID: TBSBookmarkID; out Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt8From(ID: TBSBookmarkID; out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt8From(ID: TBSBookmarkID; out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt16From(ID: TBSBookmarkID; out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt16From(ID: TBSBookmarkID; out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt32From(ID: TBSBookmarkID; out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt32From(ID: TBSBookmarkID; out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt64From(ID: TBSBookmarkID; out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt64From(ID: TBSBookmarkID; out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat32From(ID: TBSBookmarkID; out Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat64From(ID: TBSBookmarkID; out Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat80From(ID: TBSBookmarkID; out Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadDateTimeFrom(ID: TBSBookmarkID; out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCurrencyFrom(ID: TBSBookmarkID; out Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiCharFrom(ID: TBSBookmarkID; out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8CharFrom(ID: TBSBookmarkID; out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideCharFrom(ID: TBSBookmarkID; out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeCharFrom(ID: TBSBookmarkID; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4CharFrom(ID: TBSBookmarkID; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCharFrom(ID: TBSBookmarkID; out Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadShortStringFrom(ID: TBSBookmarkID; out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiStringFrom(ID: TBSBookmarkID; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8StringFrom(ID: TBSBookmarkID; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideStringFrom(ID: TBSBookmarkID; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeStringFrom(ID: TBSBookmarkID; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4StringFrom(ID: TBSBookmarkID; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadStringFrom(ID: TBSBookmarkID; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallAnsiStringFrom(ID: TBSBookmarkID; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUTF8StringFrom(ID: TBSBookmarkID; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallWideStringFrom(ID: TBSBookmarkID; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUnicodeStringFrom(ID: TBSBookmarkID; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUCS4StringFrom(ID: TBSBookmarkID; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallStringFrom(ID: TBSBookmarkID; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBufferFrom(ID: TBSBookmarkID; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadVariantFrom(ID: TBSBookmarkID; out Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFrom(ID,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.ReadBoolFromIndex(Index: Integer; out Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBooleanFromIndex(Index: Integer; out Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt8FromIndex(Index: Integer; out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt8FromIndex(Index: Integer; out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt16FromIndex(Index: Integer; out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt16FromIndex(Index: Integer; out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt32FromIndex(Index: Integer; out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt32FromIndex(Index: Integer; out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadInt64FromIndex(Index: Integer; out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUInt64FromIndex(Index: Integer; out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat32FromIndex(Index: Integer; out Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat64FromIndex(Index: Integer; out Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadFloat80FromIndex(Index: Integer; out Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadDateTimeFromIndex(Index: Integer; out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCurrencyFromIndex(Index: Integer; out Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiCharFromIndex(Index: Integer; out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8CharFromIndex(Index: Integer; out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideCharFromIndex(Index: Integer; out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeCharFromIndex(Index: Integer; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4CharFromIndex(Index: Integer; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadCharFromIndex(Index: Integer; out Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadShortStringFromIndex(Index: Integer; out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadAnsiStringFromIndex(Index: Integer; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUTF8StringFromIndex(Index: Integer; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadWideStringFromIndex(Index: Integer; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUnicodeStringFromIndex(Index: Integer; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadUCS4StringFromIndex(Index: Integer; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadStringFromIndex(Index: Integer; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallAnsiStringFromIndex(Index: Integer; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUTF8StringFromIndex(Index: Integer; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallWideStringFromIndex(Index: Integer; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUnicodeStringFromIndex(Index: Integer; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallUCS4StringFromIndex(Index: Integer; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadSmallStringFromIndex(Index: Integer; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadBufferFromIndex(Index: Integer; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.ReadVariantFromIndex(Index: Integer; out Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueFromIndex(Index,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TCustomStreamer.GetBool(Advance: Boolean = True): ByteBool;
begin
ReadValue(BS_VALTYPE_BOOL,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetBoolean(Advance: Boolean = True): Boolean;
begin
ReadValue(BS_VALTYPE_BOOLEAN,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt8(Advance: Boolean = True): Int8;
begin
ReadValue(BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt8(Advance: Boolean = True): UInt8;
begin
ReadValue(BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt16(Advance: Boolean = True): Int16;
begin
ReadValue(BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt16(Advance: Boolean = True): UInt16;
begin
ReadValue(BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt32(Advance: Boolean = True): Int32;
begin
ReadValue(BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt32(Advance: Boolean = True): UInt32;
begin
ReadValue(BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt64(Advance: Boolean = True): Int64;
begin
ReadValue(BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt64(Advance: Boolean = True): UInt64;
begin
ReadValue(BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat32(Advance: Boolean = True): Float32;
begin
ReadValue(BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat64(Advance: Boolean = True): Float64;
begin
ReadValue(BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat80(Advance: Boolean = True): Float80;
begin
ReadValue(BS_VALTYPE_NUM_10,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetDateTime(Advance: Boolean = True): TDateTime;
begin
ReadValue(BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCurrency(Advance: Boolean = True): Currency;
begin
ReadValue(BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiChar(Advance: Boolean = True): AnsiChar;
begin
ReadValue(BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8Char(Advance: Boolean = True): UTF8Char;
begin
ReadValue(BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideChar(Advance: Boolean = True): WideChar;
begin
ReadValue(BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeChar(Advance: Boolean = True): UnicodeChar;
begin
ReadValue(BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4Char(Advance: Boolean = True): UCS4Char;
begin
ReadValue(BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetChar(Advance: Boolean = True): Char;
begin
ReadValue(BS_VALTYPE_CHAR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetShortString(Advance: Boolean = True): ShortString;
begin
ReadValue(BS_VALTYPE_STR_SHORT,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiString(Advance: Boolean = True): AnsiString;
begin
ReadValue(BS_VALTYPE_STR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8String(Advance: Boolean = True): UTF8String;
begin
ReadValue(BS_VALTYPE_STR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideString(Advance: Boolean = True): WideString;
begin
ReadValue(BS_VALTYPE_STR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeString(Advance: Boolean = True): UnicodeString;
begin
ReadValue(BS_VALTYPE_STR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4String(Advance: Boolean = True): UCS4String;
begin
ReadValue(BS_VALTYPE_STR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetString(Advance: Boolean = True): String;
begin
ReadValue(BS_VALTYPE_STR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallAnsiString(Advance: Boolean = True): AnsiString;
begin
ReadValue(BS_VALTYPE_SMLSTR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUTF8String(Advance: Boolean = True): UTF8String;
begin
ReadValue(BS_VALTYPE_SMLSTR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallWideString(Advance: Boolean = True): WideString;
begin
ReadValue(BS_VALTYPE_SMLSTR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUnicodeString(Advance: Boolean = True): UnicodeString;
begin
ReadValue(BS_VALTYPE_SMLSTR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUCS4String(Advance: Boolean = True): UCS4String;
begin
ReadValue(BS_VALTYPE_SMLSTR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallString(Advance: Boolean = True): String;
begin
ReadValue(BS_VALTYPE_SMLSTR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetVariant(Advance: Boolean = True): Variant;
begin
ReadValue(BS_VALTYPE_VARIANT,@Result,Advance)
end;

//==============================================================================

Function TCustomStreamer.GetBoolAt(Position: Int64; Advance: Boolean = True): ByteBool;
begin
ReadValueAt(Position,BS_VALTYPE_BOOL,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetBooleanAt(Position: Int64; Advance: Boolean = True): Boolean;
begin
ReadValueAt(Position,BS_VALTYPE_BOOLEAN,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt8At(Position: Int64; Advance: Boolean = True): Int8;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt8At(Position: Int64; Advance: Boolean = True): UInt8;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt16At(Position: Int64; Advance: Boolean = True): Int16;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt16At(Position: Int64; Advance: Boolean = True): UInt16;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt32At(Position: Int64; Advance: Boolean = True): Int32;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt32At(Position: Int64; Advance: Boolean = True): UInt32;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt64At(Position: Int64; Advance: Boolean = True): Int64;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt64At(Position: Int64; Advance: Boolean = True): UInt64;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat32At(Position: Int64; Advance: Boolean = True): Float32;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat64At(Position: Int64; Advance: Boolean = True): Float64;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat80At(Position: Int64; Advance: Boolean = True): Float80;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_10,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetDateTimeAt(Position: Int64; Advance: Boolean = True): TDateTime;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCurrencyAt(Position: Int64; Advance: Boolean = True): Currency;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiCharAt(Position: Int64; Advance: Boolean = True): AnsiChar;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8CharAt(Position: Int64; Advance: Boolean = True): UTF8Char;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideCharAt(Position: Int64; Advance: Boolean = True): WideChar;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeCharAt(Position: Int64; Advance: Boolean = True): UnicodeChar;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4CharAt(Position: Int64; Advance: Boolean = True): UCS4Char;
begin
ReadValueAt(Position,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCharAt(Position: Int64; Advance: Boolean = True): Char;
begin
ReadValueAt(Position,BS_VALTYPE_CHAR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetShortStringAt(Position: Int64; Advance: Boolean = True): ShortString;
begin
ReadValueAt(Position,BS_VALTYPE_STR_SHORT,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiStringAt(Position: Int64; Advance: Boolean = True): AnsiString;
begin
ReadValueAt(Position,BS_VALTYPE_STR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8StringAt(Position: Int64; Advance: Boolean = True): UTF8String;
begin
ReadValueAt(Position,BS_VALTYPE_STR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideStringAt(Position: Int64; Advance: Boolean = True): WideString;
begin
ReadValueAt(Position,BS_VALTYPE_STR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeStringAt(Position: Int64; Advance: Boolean = True): UnicodeString;
begin
ReadValueAt(Position,BS_VALTYPE_STR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4StringAt(Position: Int64; Advance: Boolean = True): UCS4String;
begin
ReadValueAt(Position,BS_VALTYPE_STR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetStringAt(Position: Int64; Advance: Boolean = True): String;
begin
ReadValueAt(Position,BS_VALTYPE_STR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallAnsiStringAt(Position: Int64; Advance: Boolean = True): AnsiString;
begin
ReadValueAt(Position,BS_VALTYPE_SMLSTR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUTF8StringAt(Position: Int64; Advance: Boolean = True): UTF8String;
begin
ReadValueAt(Position,BS_VALTYPE_SMLSTR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallWideStringAt(Position: Int64; Advance: Boolean = True): WideString;
begin
ReadValueAt(Position,BS_VALTYPE_SMLSTR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUnicodeStringAt(Position: Int64; Advance: Boolean = True): UnicodeString;
begin
ReadValueAt(Position,BS_VALTYPE_SMLSTR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUCS4StringAt(Position: Int64; Advance: Boolean = True): UCS4String;
begin
ReadValueAt(Position,BS_VALTYPE_SMLSTR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallStringAt(Position: Int64; Advance: Boolean = True): String;
begin
ReadValueAt(Position,BS_VALTYPE_SMLSTR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetVariantAt(Position: Int64; Advance: Boolean = True): Variant;
begin
ReadValueAt(Position,BS_VALTYPE_VARIANT,@Result,Advance);
end;

//==============================================================================

Function TCustomStreamer.GetBoolAtOffset(Offset: Int64; Advance: Boolean = True): ByteBool;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_BOOL,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetBooleanAtOffset(Offset: Int64; Advance: Boolean = True): Boolean;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_BOOLEAN,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt8AtOffset(Offset: Int64; Advance: Boolean = True): Int8;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt8AtOffset(Offset: Int64; Advance: Boolean = True): UInt8;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt16AtOffset(Offset: Int64; Advance: Boolean = True): Int16;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt16AtOffset(Offset: Int64; Advance: Boolean = True): UInt16;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt32AtOffset(Offset: Int64; Advance: Boolean = True): Int32;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt32AtOffset(Offset: Int64; Advance: Boolean = True): UInt32;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt64AtOffset(Offset: Int64; Advance: Boolean = True): Int64;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt64AtOffset(Offset: Int64; Advance: Boolean = True): UInt64;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat32AtOffset(Offset: Int64; Advance: Boolean = True): Float32;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat64AtOffset(Offset: Int64; Advance: Boolean = True): Float64;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat80AtOffset(Offset: Int64; Advance: Boolean = True): Float80;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_10,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetDateTimeAtOffset(Offset: Int64; Advance: Boolean = True): TDateTime;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCurrencyAtOffset(Offset: Int64; Advance: Boolean = True): Currency;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiCharAtOffset(Offset: Int64; Advance: Boolean = True): AnsiChar;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8CharAtOffset(Offset: Int64; Advance: Boolean = True): UTF8Char;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideCharAtOffset(Offset: Int64; Advance: Boolean = True): WideChar;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeCharAtOffset(Offset: Int64; Advance: Boolean = True): UnicodeChar;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4CharAtOffset(Offset: Int64; Advance: Boolean = True): UCS4Char;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCharAtOffset(Offset: Int64; Advance: Boolean = True): Char;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_CHAR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetShortStringAtOffset(Offset: Int64; Advance: Boolean = True): ShortString;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_STR_SHORT,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiStringAtOffset(Offset: Int64; Advance: Boolean = True): AnsiString;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_STR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8StringAtOffset(Offset: Int64; Advance: Boolean = True): UTF8String;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_STR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideStringAtOffset(Offset: Int64; Advance: Boolean = True): WideString;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_STR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeStringAtOffset(Offset: Int64; Advance: Boolean = True): UnicodeString;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_STR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4StringAtOffset(Offset: Int64; Advance: Boolean = True): UCS4String;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_STR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetStringAtOffset(Offset: Int64; Advance: Boolean = True): String;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_STR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallAnsiStringAtOffset(Offset: Int64; Advance: Boolean = True): AnsiString;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUTF8StringAtOffset(Offset: Int64; Advance: Boolean = True): UTF8String;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallWideStringAtOffset(Offset: Int64; Advance: Boolean = True): WideString;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUnicodeStringAtOffset(Offset: Int64; Advance: Boolean = True): UnicodeString;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUCS4StringAtOffset(Offset: Int64; Advance: Boolean = True): UCS4String;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallStringAtOffset(Offset: Int64; Advance: Boolean = True): String;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_SMLSTR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetVariantAtOffset(Offset: Int64; Advance: Boolean = True): Variant;
begin
ReadValueAtOffset(Offset,BS_VALTYPE_VARIANT,@Result,Advance);
end;

//==============================================================================

Function TCustomStreamer.GetBoolFrom(ID: TBSBookmarkID; Advance: Boolean = True): ByteBool;
begin
ReadValueFrom(ID,BS_VALTYPE_BOOL,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetBooleanFrom(ID: TBSBookmarkID; Advance: Boolean = True): Boolean;
begin
ReadValueFrom(ID,BS_VALTYPE_BOOLEAN,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt8From(ID: TBSBookmarkID; Advance: Boolean = True): Int8;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt8From(ID: TBSBookmarkID; Advance: Boolean = True): UInt8;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt16From(ID: TBSBookmarkID; Advance: Boolean = True): Int16;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt16From(ID: TBSBookmarkID; Advance: Boolean = True): UInt16;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt32From(ID: TBSBookmarkID; Advance: Boolean = True): Int32;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt32From(ID: TBSBookmarkID; Advance: Boolean = True): UInt32;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt64From(ID: TBSBookmarkID; Advance: Boolean = True): Int64;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt64From(ID: TBSBookmarkID; Advance: Boolean = True): UInt64;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat32From(ID: TBSBookmarkID; Advance: Boolean = True): Float32;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat64From(ID: TBSBookmarkID; Advance: Boolean = True): Float64;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat80From(ID: TBSBookmarkID; Advance: Boolean = True): Float80;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_10,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetDateTimeFrom(ID: TBSBookmarkID; Advance: Boolean = True): TDateTime;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCurrencyFrom(ID: TBSBookmarkID; Advance: Boolean = True): Currency;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiCharFrom(ID: TBSBookmarkID; Advance: Boolean = True): AnsiChar;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8CharFrom(ID: TBSBookmarkID; Advance: Boolean = True): UTF8Char;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideCharFrom(ID: TBSBookmarkID; Advance: Boolean = True): WideChar;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeCharFrom(ID: TBSBookmarkID; Advance: Boolean = True): UnicodeChar;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4CharFrom(ID: TBSBookmarkID; Advance: Boolean = True): UCS4Char;
begin
ReadValueFrom(ID,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCharFrom(ID: TBSBookmarkID; Advance: Boolean = True): Char;
begin
ReadValueFrom(ID,BS_VALTYPE_CHAR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetShortStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): ShortString;
begin
ReadValueFrom(ID,BS_VALTYPE_STR_SHORT,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): AnsiString;
begin
ReadValueFrom(ID,BS_VALTYPE_STR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8StringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UTF8String;
begin
ReadValueFrom(ID,BS_VALTYPE_STR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): WideString;
begin
ReadValueFrom(ID,BS_VALTYPE_STR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UnicodeString;
begin
ReadValueFrom(ID,BS_VALTYPE_STR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4StringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UCS4String;
begin
ReadValueFrom(ID,BS_VALTYPE_STR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): String;
begin
ReadValueFrom(ID,BS_VALTYPE_STR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallAnsiStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): AnsiString;
begin
ReadValueFrom(ID,BS_VALTYPE_SMLSTR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUTF8StringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UTF8String;
begin
ReadValueFrom(ID,BS_VALTYPE_SMLSTR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallWideStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): WideString;
begin
ReadValueFrom(ID,BS_VALTYPE_SMLSTR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUnicodeStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UnicodeString;
begin
ReadValueFrom(ID,BS_VALTYPE_SMLSTR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUCS4StringFrom(ID: TBSBookmarkID; Advance: Boolean = True): UCS4String;
begin
ReadValueFrom(ID,BS_VALTYPE_SMLSTR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallStringFrom(ID: TBSBookmarkID; Advance: Boolean = True): String;
begin
ReadValueFrom(ID,BS_VALTYPE_SMLSTR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetVariantFrom(ID: TBSBookmarkID; Advance: Boolean = True): Variant;
begin
ReadValueFrom(ID,BS_VALTYPE_VARIANT,@Result,Advance);
end;

//==============================================================================

Function TCustomStreamer.GetBoolFromIndex(Index: Integer; Advance: Boolean = True): ByteBool;
begin
ReadValueFromIndex(Index,BS_VALTYPE_BOOL,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetBooleanFromIndex(Index: Integer; Advance: Boolean = True): Boolean;
begin
ReadValueFromIndex(Index,BS_VALTYPE_BOOLEAN,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt8FromIndex(Index: Integer; Advance: Boolean = True): Int8;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt8FromIndex(Index: Integer; Advance: Boolean = True): UInt8;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt16FromIndex(Index: Integer; Advance: Boolean = True): Int16;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt16FromIndex(Index: Integer; Advance: Boolean = True): UInt16;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt32FromIndex(Index: Integer; Advance: Boolean = True): Int32;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt32FromIndex(Index: Integer; Advance: Boolean = True): UInt32;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetInt64FromIndex(Index: Integer; Advance: Boolean = True): Int64;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUInt64FromIndex(Index: Integer; Advance: Boolean = True): UInt64;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat32FromIndex(Index: Integer; Advance: Boolean = True): Float32;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat64FromIndex(Index: Integer; Advance: Boolean = True): Float64;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetFloat80FromIndex(Index: Integer; Advance: Boolean = True): Float80;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_10,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetDateTimeFromIndex(Index: Integer; Advance: Boolean = True): TDateTime;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCurrencyFromIndex(Index: Integer; Advance: Boolean = True): Currency;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiCharFromIndex(Index: Integer; Advance: Boolean = True): AnsiChar;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8CharFromIndex(Index: Integer; Advance: Boolean = True): UTF8Char;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideCharFromIndex(Index: Integer; Advance: Boolean = True): WideChar;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeCharFromIndex(Index: Integer; Advance: Boolean = True): UnicodeChar;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4CharFromIndex(Index: Integer; Advance: Boolean = True): UCS4Char;
begin
ReadValueFromIndex(Index,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetCharFromIndex(Index: Integer; Advance: Boolean = True): Char;
begin
ReadValueFromIndex(Index,BS_VALTYPE_CHAR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetShortStringFromIndex(Index: Integer; Advance: Boolean = True): ShortString;
begin
ReadValueFromIndex(Index,BS_VALTYPE_STR_SHORT,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetAnsiStringFromIndex(Index: Integer; Advance: Boolean = True): AnsiString;
begin
ReadValueFromIndex(Index,BS_VALTYPE_STR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUTF8StringFromIndex(Index: Integer; Advance: Boolean = True): UTF8String;
begin
ReadValueFromIndex(Index,BS_VALTYPE_STR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetWideStringFromIndex(Index: Integer; Advance: Boolean = True): WideString;
begin
ReadValueFromIndex(Index,BS_VALTYPE_STR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUnicodeStringFromIndex(Index: Integer; Advance: Boolean = True): UnicodeString;
begin
ReadValueFromIndex(Index,BS_VALTYPE_STR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetUCS4StringFromIndex(Index: Integer; Advance: Boolean = True): UCS4String;
begin
ReadValueFromIndex(Index,BS_VALTYPE_STR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetStringFromIndex(Index: Integer; Advance: Boolean = True): String;
begin
ReadValueFromIndex(Index,BS_VALTYPE_STR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallAnsiStringFromIndex(Index: Integer; Advance: Boolean = True): AnsiString;
begin
ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUTF8StringFromIndex(Index: Integer; Advance: Boolean = True): UTF8String;
begin
ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallWideStringFromIndex(Index: Integer; Advance: Boolean = True): WideString;
begin
ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUnicodeStringFromIndex(Index: Integer; Advance: Boolean = True): UnicodeString;
begin
ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallUCS4StringFromIndex(Index: Integer; Advance: Boolean = True): UCS4String;
begin
ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetSmallStringFromIndex(Index: Integer; Advance: Boolean = True): String;
begin
ReadValueFromIndex(Index,BS_VALTYPE_SMLSTR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TCustomStreamer.GetVariantFromIndex(Index: Integer; Advance: Boolean = True): Variant;
begin
ReadValueFromIndex(Index,BS_VALTYPE_VARIANT,@Result,Advance);
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

Function TMemoryStreamer.GetStartAddress: Pointer;
begin
Result := PositionToAddress(fStart);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetPosition: Int64;
begin
Result := AddressToPosition(fPositionAddress);
end;

//------------------------------------------------------------------------------

procedure TMemoryStreamer.SetPosition(NewPosition: Int64);
begin
fPositionAddress := PositionToAddress(NewPosition);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
begin
case ValueType of
  BS_VALTYPE_BOOL:            Result := Ptr_WriteUInt8(fPositionAddress,UInt8(BoolToNum(ByteBool(ValuePtr^))),Advance,fEndian);
  BS_VALTYPE_BOOLEAN:         Result := Ptr_WriteUInt8(fPositionAddress,UInt8(BoolToNum(Boolean(ValuePtr^))),Advance,fEndian);
  BS_VALTYPE_NUM_1:           Result := Ptr_WriteUInt8(fPositionAddress,UInt8(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_2:           Result := Ptr_WriteUInt16(fPositionAddress,UInt16(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_4:           Result := Ptr_WriteUInt32(fPositionAddress,UInt32(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_8:           Result := Ptr_WriteUInt64(fPositionAddress,UInt64(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_10:          Result := Ptr_WriteFloat80(fPositionAddress,Float80(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_CHAR:            Result := Ptr_WriteUInt16(fPositionAddress,UInt16(Ord(Char(ValuePtr^))),Advance,fEndian);
  BS_VALTYPE_STR_SHORT:       Result := Ptr_WriteShortString(fPositionAddress,ShortString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_ANSI:        Result := Ptr_WriteAnsiString(fPositionAddress,AnsiString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UTF8:        Result := Ptr_WriteUTF8String(fPositionAddress,UTF8String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_WIDE:        Result := Ptr_WriteWideString(fPositionAddress,WideString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UNICODE:     Result := Ptr_WriteUnicodeString(fPositionAddress,UnicodeString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UCS4:        Result := Ptr_WriteUCS4String(fPositionAddress,UCS4String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR:             Result := Ptr_WriteString(fPositionAddress,String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_ANSI:     Result := Ptr_WriteSmallAnsiString(fPositionAddress,AnsiString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UTF8:     Result := Ptr_WriteSmallUTF8String(fPositionAddress,UTF8String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_WIDE:     Result := Ptr_WriteSmallWideString(fPositionAddress,WideString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UNICODE:  Result := Ptr_WriteSmallUnicodeString(fPositionAddress,UnicodeString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UCS4:     Result := Ptr_WriteSmallUCS4String(fPositionAddress,UCS4String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR:          Result := Ptr_WriteSmallString(fPositionAddress,String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_FILL:            Result := Ptr_FillBytes(fPositionAddress,Size,UInt8(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_VARIANT:         Result := Ptr_WriteVariant(fPositionAddress,Variant(ValuePtr^),Advance,fEndian);
else
 {BS_VALTYPE_BUFFER}
  Result := Ptr_WriteBuffer(fPositionAddress,ValuePtr^,Size,Advance);
end;
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteValueAtAddr(Address: Pointer; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveAt(Address);
Result := WriteValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  TempUInt8:  UInt8;
  TempUInt16: UInt16;
begin
case ValueType of
  BS_VALTYPE_BOOL:            begin
                                Result := Ptr_ReadUInt8(fPositionAddress,TempUInt8,Advance);
                                ByteBool(ValuePtr^) := NumToBool(TempUInt8);
                              end;
  BS_VALTYPE_BOOLEAN:         begin
                                Result := Ptr_ReadUInt8(fPositionAddress,TempUInt8,Advance);
                                Boolean(ValuePtr^) := NumToBool(TempUInt8);
                              end;
  BS_VALTYPE_NUM_1:           Result := Ptr_ReadUInt8(fPositionAddress,UInt8(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_2:           Result := Ptr_ReadUInt16(fPositionAddress,UInt16(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_4:           Result := Ptr_ReadUInt32(fPositionAddress,UInt32(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_8:           Result := Ptr_ReadUInt64(fPositionAddress,UInt64(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_10:          Result := Ptr_ReadFloat80(fPositionAddress,Float80(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_CHAR:            begin
                                Result := Ptr_ReadUInt16(fPositionAddress,TempUInt16,Advance);
                                Char(ValuePtr^) := Char(TempUInt16);
                              end;
  BS_VALTYPE_STR_SHORT:       Result := Ptr_ReadShortString(fPositionAddress,ShortString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_ANSI:        Result := Ptr_ReadAnsiString(fPositionAddress,AnsiString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UTF8:        Result := Ptr_ReadUTF8String(fPositionAddress,UTF8String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_WIDE:        Result := Ptr_ReadWideString(fPositionAddress,WideString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UNICODE:     Result := Ptr_ReadUnicodeString(fPositionAddress,UnicodeString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UCS4:        Result := Ptr_ReadUCS4String(fPositionAddress,UCS4String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR:             Result := Ptr_ReadString(fPositionAddress,String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_ANSI:     Result := Ptr_ReadSmallAnsiString(fPositionAddress,AnsiString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UTF8:     Result := Ptr_ReadSmallUTF8String(fPositionAddress,UTF8String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_WIDE:     Result := Ptr_ReadSmallWideString(fPositionAddress,WideString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UNICODE:  Result := Ptr_ReadSmallUnicodeString(fPositionAddress,UnicodeString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UCS4:     Result := Ptr_ReadSmallUCS4String(fPositionAddress,UCS4String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR:          Result := Ptr_ReadSmallString(fPositionAddress,String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_VARIANT:         Result := Ptr_ReadVariant(fPositionAddress,Variant(ValuePtr^),Advance,fEndian);
else
 {BS_VALTYPE_FILL,BS_VALTYPE_BUFFER}
  Result := Ptr_ReadBuffer(fPositionAddress,ValuePtr^,Size,Advance);
end;
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadValueAtAddr(Address: Pointer; ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  OldPosition:  Int64;
begin
OldPosition := Self.Position;
MoveAt(Address);
Result := ReadValue(ValueType,ValuePtr,Advance,Size);
If not Advance then
  Self.Position := OldPosition;
end;

//------------------------------------------------------------------------------

procedure TMemoryStreamer.Initialize(Memory: Pointer);
begin
inherited Initialize;
fStart := AddressToPosition(Memory);
fPositionAddress := Memory;
end;

{-------------------------------------------------------------------------------
    TMemoryStreamer - public methods
-------------------------------------------------------------------------------}

class Function TMemoryStreamer.AddressToPosition(Address: Pointer): Int64;
begin
{$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
Result := Int64(PtrUInt(Address));
{$IFDEF FPCDWM}{$POP}{$ENDIF}
end;

//------------------------------------------------------------------------------

class Function TMemoryStreamer.PositionToAddress(Position: Int64): Pointer;
begin
{$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
Result := Pointer(PtrUInt(Position));
{$IFDEF FPCDWM}{$POP}{$ENDIF}
end;

//------------------------------------------------------------------------------

constructor TMemoryStreamer.Create(Memory: Pointer);
begin
inherited Create;
Initialize(Memory);
end;

//------------------------------------------------------------------------------

procedure TMemoryStreamer.MoveAt(Address: Pointer);
begin
fPositionAddress := Address;
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.BookmarkAdd(ID: TBSBookmarkID; Address: Pointer): Integer;
begin
Result := BookmarkAdd(ID,AddressToPosition(Address));
end;

//------------------------------------------------------------------------------

procedure TMemoryStreamer.BookmarkInsert(Index: Integer; ID: TBSBookmarkID; Address: Pointer);
begin
BookmarkInsert(Index,ID,AddressToPosition(Address));
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.BookmarkGetAddress(ID: TBSBookmarkID): Pointer;
begin
Result := PositionToAddress(BookmarkGetPosition(ID));
end;

//------------------------------------------------------------------------------

procedure TMemoryStreamer.BookmarkSetAddress(ID: TBSBookmarkID; NewAddress: Pointer);
begin
BookmarkSetPosition(ID,AddressToPosition(NewAddress));
end;

//==============================================================================

Function TMemoryStreamer.WriteBoolAt(Address: Pointer; Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteBooleanAt(Address: Pointer; Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteInt8At(Address: Pointer; Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUInt8At(Address: Pointer; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteInt16At(Address: Pointer; Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUInt16At(Address: Pointer; Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteInt32At(Address: Pointer; Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUInt32At(Address: Pointer; Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteInt64At(Address: Pointer; Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUInt64At(Address: Pointer; Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteFloat32At(Address: Pointer; Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteFloat64At(Address: Pointer; Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteFloat80At(Address: Pointer; Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteDateTimeAt(Address: Pointer; Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteCurrencyAt(Address: Pointer; Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteAnsiCharAt(Address: Pointer; Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUTF8CharAt(Address: Pointer; Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteWideCharAt(Address: Pointer; Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUnicodeCharAt(Address: Pointer; Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUCS4CharAt(Address: Pointer; Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteCharAt(Address: Pointer; Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteShortStringAt(Address: Pointer; const Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteAnsiStringAt(Address: Pointer; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUTF8StringAt(Address: Pointer; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteWideStringAt(Address: Pointer; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUnicodeStringAt(Address: Pointer; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteUCS4StringAt(Address: Pointer; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteStringAt(Address: Pointer; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteSmallAnsiStringAt(Address: Pointer; const Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteSmallUTF8StringAt(Address: Pointer; const Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteSmallWideStringAt(Address: Pointer; const Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteSmallUnicodeStringAt(Address: Pointer; const Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteSmallUCS4StringAt(Address: Pointer; const Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteSmallStringAt(Address: Pointer; const Value: String; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteBufferAt(Address: Pointer; const Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteBytesAt(Address: Pointer; const Value: array of UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_BYTEARRAY,@Value,Advance,TMemSize(Length(Value)));
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.FillBytesAt(Address: Pointer; Count: TMemSize; Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_FILL,@Value,Advance,Count);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.WriteVariantAt(Address: Pointer; const Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := WriteValueAtAddr(Address,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TMemoryStreamer.ReadBoolAt(Address: Pointer; out Value: ByteBool; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_BOOL,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadBooleanAt(Address: Pointer; out Value: Boolean; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_BOOLEAN,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadInt8At(Address: Pointer; out Value: Int8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUInt8At(Address: Pointer; out Value: UInt8; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadInt16At(Address: Pointer; out Value: Int16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUInt16At(Address: Pointer; out Value: UInt16; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadInt32At(Address: Pointer; out Value: Int32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUInt32At(Address: Pointer; out Value: UInt32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadInt64At(Address: Pointer; out Value: Int64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUInt64At(Address: Pointer; out Value: UInt64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadFloat32At(Address: Pointer; out Value: Float32; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadFloat64At(Address: Pointer; out Value: Float64; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadFloat80At(Address: Pointer; out Value: Float80; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_10,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadDateTimeAt(Address: Pointer; out Value: TDateTime; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadCurrencyAt(Address: Pointer; out Value: Currency; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadAnsiCharAt(Address: Pointer; out Value: AnsiChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUTF8CharAt(Address: Pointer; out Value: UTF8Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_1,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadWideCharAt(Address: Pointer; out Value: WideChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUnicodeCharAt(Address: Pointer; out Value: UnicodeChar; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_2,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUCS4CharAt(Address: Pointer; out Value: UCS4Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_NUM_4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadCharAt(Address: Pointer; out Value: Char; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_CHAR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadShortStringAt(Address: Pointer; out Value: ShortString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_STR_SHORT,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadAnsiStringAt(Address: Pointer; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_STR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUTF8StringAt(Address: Pointer; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_STR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadWideStringAt(Address: Pointer; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_STR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUnicodeStringAt(Address: Pointer; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_STR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadUCS4StringAt(Address: Pointer; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_STR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadStringAt(Address: Pointer; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_STR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadSmallAnsiStringAt(Address: Pointer; out Value: AnsiString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_ANSI,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadSmallUTF8StringAt(Address: Pointer; out Value: UTF8String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_UTF8,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadSmallWideStringAt(Address: Pointer; out Value: WideString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_WIDE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadSmallUnicodeStringAt(Address: Pointer; out Value: UnicodeString; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_UNICODE,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadSmallUCS4StringAt(Address: Pointer; out Value: UCS4String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_UCS4,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadSmallStringAt(Address: Pointer; out Value: String; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR,@Value,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadBufferAt(Address: Pointer; out Buffer; Size: TMemSize; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_BUFFER,@Buffer,Advance,Size);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.ReadVariantAt(Address: Pointer; out Value: Variant; Advance: Boolean = True): TMemSize;
begin
Result := ReadValueAtAddr(Address,BS_VALTYPE_VARIANT,@Value,Advance);
end;

//==============================================================================

Function TMemoryStreamer.GetBoolAt(Address: Pointer; Advance: Boolean = True): ByteBool;
begin
ReadValueAtAddr(Address,BS_VALTYPE_BOOL,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetBooleanAt(Address: Pointer; Advance: Boolean = True): Boolean;
begin
ReadValueAtAddr(Address,BS_VALTYPE_BOOLEAN,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetInt8At(Address: Pointer; Advance: Boolean = True): Int8;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUInt8At(Address: Pointer; Advance: Boolean = True): UInt8;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetInt16At(Address: Pointer; Advance: Boolean = True): Int16;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUInt16At(Address: Pointer; Advance: Boolean = True): UInt16;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetInt32At(Address: Pointer; Advance: Boolean = True): Int32;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUInt32At(Address: Pointer; Advance: Boolean = True): UInt32;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetInt64At(Address: Pointer; Advance: Boolean = True): Int64;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUInt64At(Address: Pointer; Advance: Boolean = True): UInt64;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetFloat32At(Address: Pointer; Advance: Boolean = True): Float32;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetFloat64At(Address: Pointer; Advance: Boolean = True): Float64;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetFloat80At(Address: Pointer; Advance: Boolean = True): Float80;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_10,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetDateTimeAt(Address: Pointer; Advance: Boolean = True): TDateTime;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetCurrencyAt(Address: Pointer; Advance: Boolean = True): Currency;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetAnsiCharAt(Address: Pointer; Advance: Boolean = True): AnsiChar;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUTF8CharAt(Address: Pointer; Advance: Boolean = True): UTF8Char;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_1,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetWideCharAt(Address: Pointer; Advance: Boolean = True): WideChar;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUnicodeCharAt(Address: Pointer; Advance: Boolean = True): UnicodeChar;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_2,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUCS4CharAt(Address: Pointer; Advance: Boolean = True): UCS4Char;
begin
ReadValueAtAddr(Address,BS_VALTYPE_NUM_4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetCharAt(Address: Pointer; Advance: Boolean = True): Char;
begin
ReadValueAtAddr(Address,BS_VALTYPE_CHAR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetShortStringAt(Address: Pointer; Advance: Boolean = True): ShortString;
begin
ReadValueAtAddr(Address,BS_VALTYPE_STR_SHORT,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetAnsiStringAt(Address: Pointer; Advance: Boolean = True): AnsiString;
begin
ReadValueAtAddr(Address,BS_VALTYPE_STR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUTF8StringAt(Address: Pointer; Advance: Boolean = True): UTF8String;
begin
ReadValueAtAddr(Address,BS_VALTYPE_STR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetWideStringAt(Address: Pointer; Advance: Boolean = True): WideString;
begin
ReadValueAtAddr(Address,BS_VALTYPE_STR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUnicodeStringAt(Address: Pointer; Advance: Boolean = True): UnicodeString;
begin
ReadValueAtAddr(Address,BS_VALTYPE_STR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetUCS4StringAt(Address: Pointer; Advance: Boolean = True): UCS4String;
begin
ReadValueAtAddr(Address,BS_VALTYPE_STR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetStringAt(Address: Pointer; Advance: Boolean = True): String;
begin
ReadValueAtAddr(Address,BS_VALTYPE_STR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetSmallAnsiStringAt(Address: Pointer; Advance: Boolean = True): AnsiString;
begin
ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_ANSI,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetSmallUTF8StringAt(Address: Pointer; Advance: Boolean = True): UTF8String;
begin
ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_UTF8,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetSmallWideStringAt(Address: Pointer; Advance: Boolean = True): WideString;
begin
ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_WIDE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetSmallUnicodeStringAt(Address: Pointer; Advance: Boolean = True): UnicodeString;
begin
ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_UNICODE,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetSmallUCS4StringAt(Address: Pointer; Advance: Boolean = True): UCS4String;
begin
ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR_UCS4,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetSmallStringAt(Address: Pointer; Advance: Boolean = True): String;
begin
ReadValueAtAddr(Address,BS_VALTYPE_SMLSTR,@Result,Advance);
end;

//------------------------------------------------------------------------------

Function TMemoryStreamer.GetVariantAt(Address: Pointer; Advance: Boolean = True): Variant;
begin
ReadValueAtAddr(Address,BS_VALTYPE_VARIANT,@Result,Advance);
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

Function TStreamStreamer.GetPosition: Int64;
begin
Result := fTarget.Position;
end;

//------------------------------------------------------------------------------

procedure TStreamStreamer.SetPosition(NewPosition: Int64);
begin
fTarget.Position := NewPosition;
end;

//------------------------------------------------------------------------------

Function TStreamStreamer.WriteValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
begin
case ValueType of
  BS_VALTYPE_BOOL:            Result := Stream_WriteUInt8(fTarget,UInt8(BoolToNum(ByteBool(ValuePtr^))),Advance,fEndian);
  BS_VALTYPE_BOOLEAN:         Result := Stream_WriteUInt8(fTarget,UInt8(BoolToNum(Boolean(ValuePtr^))),Advance,fEndian);
  BS_VALTYPE_NUM_1:           Result := Stream_WriteUInt8(fTarget,UInt8(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_2:           Result := Stream_WriteUInt16(fTarget,UInt16(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_4:           Result := Stream_WriteUInt32(fTarget,UInt32(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_8:           Result := Stream_WriteUInt64(fTarget,UInt64(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_NUM_10:          Result := Stream_WriteFloat80(fTarget,Float80(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_CHAR:            Result := Stream_WriteUInt16(fTarget,UInt16(Ord(Char(ValuePtr^))),Advance,fEndian);
  BS_VALTYPE_STR_SHORT:       Result := Stream_WriteShortString(fTarget,ShortString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_ANSI:        Result := Stream_WriteAnsiString(fTarget,AnsiString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UTF8:        Result := Stream_WriteUTF8String(fTarget,UTF8String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_WIDE:        Result := Stream_WriteWideString(fTarget,WideString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UNICODE:     Result := Stream_WriteUnicodeString(fTarget,UnicodeString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR_UCS4:        Result := Stream_WriteUCS4String(fTarget,UCS4String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_STR:             Result := Stream_WriteString(fTarget,String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_ANSI:     Result := Stream_WriteSmallAnsiString(fTarget,AnsiString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UTF8:     Result := Stream_WriteSmallUTF8String(fTarget,UTF8String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_WIDE:     Result := Stream_WriteSmallWideString(fTarget,WideString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UNICODE:  Result := Stream_WriteSmallUnicodeString(fTarget,UnicodeString(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR_UCS4:     Result := Stream_WriteSmallUCS4String(fTarget,UCS4String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_SMLSTR:          Result := Stream_WriteSmallString(fTarget,String(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_BYTEARRAY:       Result := WriteOpenByteArray(ValuePtr,Size,Advance);
  BS_VALTYPE_FILL:            Result := Stream_FillBytes(fTarget,Size,UInt8(ValuePtr^),Advance,fEndian);
  BS_VALTYPE_VARIANT:         Result := Stream_WriteVariant(fTarget,Variant(ValuePtr^),Advance,fEndian);
else
 {BS_VALTYPE_BUFFER}
  Result := Stream_WriteBuffer(fTarget,ValuePtr^,Size,Advance,fEndian);
end;
end;

//------------------------------------------------------------------------------

Function TStreamStreamer.ReadValue(ValueType: Integer; ValuePtr: Pointer; Advance: Boolean; Size: TMemSize = 0): TMemSize;
var
  TempUInt8:  UInt8;
  TempUInt16: UInt16;
begin
case ValueType of
  BS_VALTYPE_BOOL:            begin
                                Result := Stream_ReadUInt8(fTarget,TempUInt8,Advance);
                                ByteBool(ValuePtr^) := NumToBool(TempUInt8);
                              end;
  BS_VALTYPE_BOOLEAN:         begin
                                Result := Stream_ReadUInt8(fTarget,TempUInt8,Advance);
                                Boolean(ValuePtr^) := NumToBool(TempUInt8);
                              end;
  BS_VALTYPE_NUM_1:           Result := Stream_ReadUInt8(fTarget,UInt8(ValuePtr^),Advance);
  BS_VALTYPE_NUM_2:           Result := Stream_ReadUInt16(fTarget,UInt16(ValuePtr^),Advance);
  BS_VALTYPE_NUM_4:           Result := Stream_ReadUInt32(fTarget,UInt32(ValuePtr^),Advance);
  BS_VALTYPE_NUM_8:           Result := Stream_ReadUInt64(fTarget,UInt64(ValuePtr^),Advance);
  BS_VALTYPE_NUM_10:          Result := Stream_ReadFloat80(fTarget,Float80(ValuePtr^),Advance);
  BS_VALTYPE_CHAR:            begin
                                Result := Stream_ReadUInt16(fTarget,TempUInt16,Advance);
                                Char(ValuePtr^) := Char(TempUInt16);
                              end;
  BS_VALTYPE_STR_SHORT:       Result := Stream_ReadShortString(fTarget,ShortString(ValuePtr^),Advance);
  BS_VALTYPE_STR_ANSI:        Result := Stream_ReadAnsiString(fTarget,AnsiString(ValuePtr^),Advance);
  BS_VALTYPE_STR_UTF8:        Result := Stream_ReadUTF8String(fTarget,UTF8String(ValuePtr^),Advance);
  BS_VALTYPE_STR_WIDE:        Result := Stream_ReadWideString(fTarget,WideString(ValuePtr^),Advance);
  BS_VALTYPE_STR_UNICODE:     Result := Stream_ReadUnicodeString(fTarget,UnicodeString(ValuePtr^),Advance);
  BS_VALTYPE_STR_UCS4:        Result := Stream_ReadUCS4String(fTarget,UCS4String(ValuePtr^),Advance);
  BS_VALTYPE_STR:             Result := Stream_ReadString(fTarget,String(ValuePtr^),Advance);
  BS_VALTYPE_SMLSTR_ANSI:     Result := Stream_ReadSmallAnsiString(fTarget,AnsiString(ValuePtr^),Advance);
  BS_VALTYPE_SMLSTR_UTF8:     Result := Stream_ReadSmallUTF8String(fTarget,UTF8String(ValuePtr^),Advance);
  BS_VALTYPE_SMLSTR_WIDE:     Result := Stream_ReadSmallWideString(fTarget,WideString(ValuePtr^),Advance);
  BS_VALTYPE_SMLSTR_UNICODE:  Result := Stream_ReadSmallUnicodeString(fTarget,UnicodeString(ValuePtr^),Advance);
  BS_VALTYPE_SMLSTR_UCS4:     Result := Stream_ReadSmallUCS4String(fTarget,UCS4String(ValuePtr^),Advance);
  BS_VALTYPE_SMLSTR:          Result := Stream_ReadSmallString(fTarget,String(ValuePtr^),Advance);
  BS_VALTYPE_VARIANT:         Result := Stream_ReadVariant(fTarget,Variant(ValuePtr^),Advance);
else
 {BS_VALTYPE_BUFFER,BS_VALTYPE_BYTEARRAY,BS_VALTYPE_FILL}
  Result := Stream_ReadBuffer(fTarget,ValuePtr^,Size,Advance);
end;
end;

//------------------------------------------------------------------------------

procedure TStreamStreamer.Initialize(Target: TStream);
begin
inherited Initialize;
fStart := Target.Position;
fTarget := Target;
end;

{-------------------------------------------------------------------------------
    TStreamStreamer - public methods
-------------------------------------------------------------------------------}

constructor TStreamStreamer.Create(Target: TStream);
begin
inherited Create;
Initialize(Target);
end;


{<lite-begin>}
{===============================================================================
--------------------------------------------------------------------------------
                                Unit preparation
--------------------------------------------------------------------------------
===============================================================================}

procedure UnitInitialize;

  Function GetOBAStride(const Arr: array of UInt8): Integer;
  begin
    If Length(Arr) >= 2 then
    {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
      Result := Integer(PtrUInt(Addr(Arr[1])) - PtrUInt(Addr(Arr[0])))
    {$IFDEF FPCDWM}{$POP}{$ENDIF}
    else
      Result := 0;
  end;

begin
OpenByteArrayStride := GetOBAStride([0,0]);
// following should be always true, but the paranoia needs feeding...
OpenByteArrayIsPacked := OpenByteArrayStride = 1;
end;

//==============================================================================
{<lite-end-ln>}

initialization
  {<lite-initialization>}
  {<lite-begin>}
  UnitInitialize;
  {<lite-end-ln>}

end.

