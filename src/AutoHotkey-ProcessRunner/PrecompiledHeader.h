// WinApi
#include <windows.h>
#include <winscard.h>

// ATL
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
#include <atlsafe.h>
#include <atlstr.h>
#include <comdef.h>

// std
#include <filesystem>
#include <ranges>

// Boost
#include <boost/process/v2/process.hpp>
#include <boost/process/v2/windows/show_window.hpp>
#include <boost/process/v2/stdio.hpp>
#include <boost/process/v2/environment.hpp>
#include <boost/process/v2/start_dir.hpp>
#include <boost/asio.hpp>
#include <boost/noncopyable.hpp>
#include <boost/locale.hpp>
#include <boost/filesystem.hpp>
