#include <AutoHotkey-ProcessRunner/Gen/ProcessRunner_i.h>

class ATL_NO_VTABLE ProcessRunnerResult :
  public CComObjectRootEx<CComSingleThreadModel>,
  public CComCoClass<ProcessRunnerResult, &__uuidof(ProcessRunnerResult)>,
  public IDispatchImpl<IProcessRunnerResult, &IID_IProcessRunnerResult, &LIBID_AutoHotkeyProcessRunnerLib, /* wMajor */ 0xFFFF, /* wMinor */ 0xFFFF> {
public:
  BEGIN_COM_MAP(ProcessRunnerResult)
    COM_INTERFACE_ENTRY(IProcessRunnerResult)
    COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()

  STDMETHOD(get_StdOut)(BSTR* val) override try {
    *val = Copy(_stdOut);
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_StdErr)(BSTR* val) override try {
    *val = Copy(_stdErr);
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_ExitCode)(long* val) override try {
    *val = _exitCode;
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  void Init(std::wstring stdOut, std::wstring stdErr, const long exitCode) {
    _stdOut = std::move(stdOut);
    _stdErr = std::move(stdErr);
    _exitCode = exitCode;
  }

private:
  std::wstring _stdOut;
  std::wstring _stdErr;
  long _exitCode;
};

class ATL_NO_VTABLE ProcessRunner :
  public CComObjectRootEx<CComSingleThreadModel>,
  public CComCoClass<ProcessRunner, &CLSID_ProcessRunner>,
  public IDispatchImpl<IProcessRunner, &IID_IProcessRunner, &LIBID_AutoHotkeyProcessRunnerLib, /* wMajor */ 0xFFFF, /* wMinor */ 0xFFFF> {
public:
  BEGIN_COM_MAP(ProcessRunner)
    COM_INTERFACE_ENTRY(IProcessRunner)
    COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()

  STDMETHOD(Run)(BSTR executablePath, IDispatch* commandLineArgumentsAutoHotkeyArray, BSTR workingDirectory, IProcessRunnerResult** result) override try {
    auto [stdOut, stdErr, exitCode] = RunProcess(
      ExpandPathWithEnvironmentVariables(executablePath),
      ToUtf8StringVector(AutoHotkeyStringArrayToVector(commandLineArgumentsAutoHotkeyArray)),
      ExpandPathWithEnvironmentVariables(workingDirectory));
    *result = CreateComObject<ProcessRunnerResult, IProcessRunnerResult>(
      [&](auto& obj) {
        obj.Init(std::move(stdOut), std::move(stdErr), exitCode);
      }).Detach();
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

private:
  static std::tuple<std::wstring, std::wstring, int> RunProcess(const std::filesystem::path& exePath,
                                                                const std::vector<std::string>& args,
                                                                const std::filesystem::path& workingDirectory) {
    if (!std::filesystem::exists(exePath)) {
      THROW_WEXCEPTION(L"The executable not found '{}'", exePath);
    }

    if (!boost::process::v2::environment::detail::is_exec_type(exePath.c_str())) {
      THROW_WEXCEPTION(L"The file '{}' is not executable", exePath);
    }

    boost::asio::io_context ctx;
    boost::asio::streambuf outBuffer, errBuffer;
    boost::asio::readable_pipe stdOutPipe(ctx), stdErrPipe(ctx);
    boost::process::v2::process proc(
      ctx, std::filesystem::canonical(exePath), args,
      boost::process::v2::process_stdio{ nullptr, stdOutPipe, stdErrPipe },
      boost::process::v2::windows::show_window_hide,
      boost::process::v2::process_start_dir(workingDirectory));

    std::function<void()> readStdOut;
    readStdOut = [&]() {
      boost::asio::async_read(
        stdOutPipe, outBuffer,
        [&](const boost::system::error_code& ec, std::size_t /* bytes_transferred */) {
          if (!ec) {
            readStdOut();
          }
        });
      };

    std::function<void()> readStdErr;
    readStdErr = [&]() {
      boost::asio::async_read(
        stdErrPipe, errBuffer,
        [&](const boost::system::error_code& ec, std::size_t /* bytes_transferred */) {
          if (!ec) {
            readStdErr();
          }
        });
      };

    readStdOut();
    readStdErr();

    ctx.run();

    const auto& exitCode = proc.wait();

    return { ToUtf16({ GetDataPointer(outBuffer), outBuffer.size() }),
             ToUtf16({ GetDataPointer(errBuffer), errBuffer.size() }),
             exitCode };
  }
};
