#pragma once

const wchar_t* VarTypeToString(const VARTYPE type) {
  switch (type) {
  case VT_EMPTY: return L"VT_EMPTY";
  case VT_NULL: return L"VT_NULL";
  case VT_I2: return L"VT_I2";
  case VT_I4: return L"VT_I4";
  case VT_R4: return L"VT_R4";
  case VT_R8: return L"VT_R8";
  case VT_CY: return L"VT_CY";
  case VT_DATE: return L"VT_DATE";
  case VT_BSTR: return L"VT_BSTR";
  case VT_DISPATCH: return L"VT_DISPATCH";
  case VT_ERROR: return L"VT_ERROR";
  case VT_BOOL: return L"VT_BOOL";
  case VT_VARIANT: return L"VT_VARIANT";
  case VT_UNKNOWN: return L"VT_UNKNOWN";
  case VT_DECIMAL: return L"VT_DECIMAL";
  case VT_I1: return L"VT_I1";
  case VT_UI1: return L"VT_UI1";
  case VT_UI2: return L"VT_UI2";
  case VT_UI4: return L"VT_UI4";
  case VT_I8: return L"VT_I8";
  case VT_UI8: return L"VT_UI8";
  case VT_INT: return L"VT_INT";
  case VT_UINT: return L"VT_UINT";
  case VT_VOID: return L"VT_VOID";
  case VT_HRESULT: return L"VT_HRESULT";
  case VT_PTR: return L"VT_PTR";
  case VT_SAFEARRAY: return L"VT_SAFEARRAY";
  case VT_CARRAY: return L"VT_CARRAY";
  case VT_USERDEFINED: return L"VT_USERDEFINED";
  case VT_LPSTR: return L"VT_LPSTR";
  case VT_LPWSTR: return L"VT_LPWSTR";
  case VT_RECORD: return L"VT_RECORD";
  case VT_INT_PTR: return L"VT_INT_PTR";
  case VT_UINT_PTR: return L"VT_UINT_PTR";
  case VT_FILETIME: return L"VT_FILETIME";
  case VT_BLOB: return L"VT_BLOB";
  case VT_STREAM: return L"VT_STREAM";
  case VT_STORAGE: return L"VT_STORAGE";
  case VT_STREAMED_OBJECT: return L"VT_STREAMED_OBJECT";
  case VT_STORED_OBJECT: return L"VT_STORED_OBJECT";
  case VT_BLOB_OBJECT: return L"VT_BLOB_OBJECT";
  case VT_CF: return L"VT_CF";
  case VT_CLSID: return L"VT_CLSID";
  case VT_VERSIONED_STREAM: return L"VT_VERSIONED_STREAM";
  default: return L"Unknown VARTYPE";
  }
}
