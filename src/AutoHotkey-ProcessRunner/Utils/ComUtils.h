#pragma once

inline DISPID GetPropertyOrMethodId(IDispatch& obj, const std::wstring_view propOrMethodName) {
  DISPID dispId = 0;
  LPOLESTR lpNames[] = { const_cast<LPOLESTR>(propOrMethodName.data()) };
  THROW_IF_FAILED_MSG(
    obj.GetIDsOfNames(IID_NULL, lpNames, 1, LOCALE_USER_DEFAULT, &dispId),
    L"IDispatch object doesn't have a property or method '{}'", propOrMethodName);
  return dispId;
}

inline CComVariant Invoke(
  IDispatch& obj,
  const std::wstring_view methodOrPropertyName,
  VARIANT* args = nullptr, // args must be in reverse order
  UINT argCount = 0,
  int invokeType = DISPATCH_PROPERTYGET
) {
  const auto& dispId = GetPropertyOrMethodId(obj, methodOrPropertyName);

  CComVariant result;
  DISPPARAMS dispParams{};
  dispParams.cArgs = argCount;
  dispParams.rgvarg = args;

  THROW_IF_FAILED_MSG(
    obj.Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, invokeType, &dispParams, &result, nullptr, nullptr),
    L"Failed to get the value of the property or call a method '{}'", methodOrPropertyName);

  return result;
}

inline std::vector<std::wstring> AutoHotkeyStringArrayToVector(IDispatch* obj) {
  std::vector<std::wstring> result;

  if (obj == 0) {
    return result;
  }

  const auto& lengthVariant = Invoke(*obj, L"Length");
  if (lengthVariant.vt != VT_I4) {
      THROW_WEXCEPTION(L"The property 'Length' of the array has a wrong variant type {}. Expected type is VT_I4", VarTypeToString(lengthVariant.vt));
  }

  // AutoHotkey arrays are 1-based
  for (LONG i = 1; i <= lengthVariant.lVal; i++) {
    // Execute ArrayObj.Get(Index , Default)
    CComVariant defaultVal(L"");
    CComVariant indexVal(i);      
    VARIANT args[2] = { defaultVal, indexVal }; // arguments in reverse order

    const auto value = Invoke(*obj, L"Get", args, 2, DISPATCH_METHOD);
    if (value.vt != VT_BSTR) {
      THROW_WEXCEPTION(L"The element with the index {} is not a string but a variant type {}", i, VarTypeToString(value.vt));
    }
    result.emplace_back(value.bstrVal);
  }

  return result;
}

template<typename CoClassType, typename QueryInterfaceType>
CComPtr<QueryInterfaceType> CreateComObject(std::function<void(CoClassType&)> initializer = [](auto&) {}) {
  CComObject<CoClassType>* rawObj = nullptr;
  THROW_IF_FAILED_MSG(
    CComObject<CoClassType>::CreateInstance(&rawObj),
    L"Failed to create a COM object of type '{}'", ToUtf16(typeid(CoClassType).name()));

  // Wrap the object in CComPtr to ensure that it is released in case of an exception
  CComPtr<CoClassType> obj(rawObj);

  initializer(*obj);

  CComQIPtr<QueryInterfaceType> res(obj);
  if (!res) {
    THROW_HRESULT(
      E_NOINTERFACE,
      L"Failed to get the interface '{}' from '{}' class", ToUtf16(typeid(QueryInterfaceType).name()), ToUtf16(typeid(CoClassType).name()));
  }

  return res;
}

inline BSTR Copy(std::wstring_view str) {
  const auto& res = SysAllocStringLen(str.data(), static_cast<UINT>(str.size()));
  if (res == nullptr) {
    THROW_WEXCEPTION(L"Failed to create a BSTR copy");
  }
  return res;
}
