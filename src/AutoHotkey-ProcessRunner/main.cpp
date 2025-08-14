#include <AutoHotkey-ProcessRunner/Utils/StringUtils.h>
#include <AutoHotkey-ProcessRunner/Utils/Logger.h>
#include <AutoHotkey-ProcessRunner/Utils/Exception.h>
#include <AutoHotkey-ProcessRunner/Utils/WinApiUtils.h>
#include <AutoHotkey-ProcessRunner/Utils/ComVariantToString.h>
#include <AutoHotkey-ProcessRunner/Utils/ComUtils.h>
#include <AutoHotkey-ProcessRunner/ProcessRunner.h>

class CProcessRunnerModule : public ATL::CAtlDllModuleT<CProcessRunnerModule> {
public:
  DECLARE_LIBID(LIBID_AutoHotkeyProcessRunnerLib)
};
CProcessRunnerModule _AtlModule;

extern "C" BOOL WINAPI DllMain(const HINSTANCE /* thisDllModule */, const DWORD reason, const LPVOID reserved) {
  return _AtlModule.DllMain(reason, reserved);
}

extern "C" __declspec(dllexport) IProcessRunner* WINAPI CreateProcessRunner() try {
  return CreateComObject<ProcessRunner, IProcessRunner>().Detach();
} catch (const std::exception&) {
  return nullptr;
}
