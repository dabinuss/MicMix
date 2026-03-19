#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <array>
#include <cstring>
#include <string>

namespace {

std::wstring GetArgValue(int argc, wchar_t** argv, const wchar_t* key, const wchar_t* fallback) {
    if (!argv || argc <= 0) {
        return fallback ? fallback : L"";
    }
    for (int i = 1; i < argc; ++i) {
        if (_wcsicmp(argv[i], key) == 0 && (i + 1) < argc) {
            return argv[i + 1];
        }
    }
    return fallback ? fallback : L"";
}

std::wstring BuildPipePath(const std::wstring& pipeName) {
    std::wstring name = pipeName;
    if (name.empty()) {
        name = L"micmix_vst_host";
    }
    name.erase(std::remove_if(name.begin(), name.end(), [](wchar_t ch) {
        return ch == L'\\' || ch == L'/' || ch == L':';
    }), name.end());
    if (name.empty()) {
        name = L"micmix_vst_host";
    }
    return L"\\\\.\\pipe\\" + name;
}

bool ReadCommand(HANDLE pipe, std::string& outLine) {
    outLine.clear();
    std::array<char, 256> buffer{};
    DWORD readBytes = 0;
    if (!ReadFile(pipe, buffer.data(), static_cast<DWORD>(buffer.size() - 1), &readBytes, nullptr) || readBytes == 0) {
        return false;
    }
    buffer[readBytes] = '\0';
    outLine.assign(buffer.data(), buffer.data() + readBytes);
    while (!outLine.empty() && (outLine.back() == '\r' || outLine.back() == '\n' || outLine.back() == '\0')) {
        outLine.pop_back();
    }
    return true;
}

void WriteResponse(HANDLE pipe, const char* text) {
    if (!text) {
        return;
    }
    DWORD written = 0;
    const DWORD len = static_cast<DWORD>(std::strlen(text));
    WriteFile(pipe, text, len, &written, nullptr);
}

} // namespace

int wmain(int argc, wchar_t** argv) {
    const std::wstring pipeName = GetArgValue(argc, argv, L"--pipe", L"micmix_vst_host");
    const std::wstring pipePath = BuildPipePath(pipeName);

    std::atomic<bool> running{true};
    while (running.load(std::memory_order_acquire)) {
        HANDLE pipe = CreateNamedPipeW(
            pipePath.c_str(),
            PIPE_ACCESS_DUPLEX,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
            1,
            4096,
            4096,
            200,
            nullptr);
        if (pipe == INVALID_HANDLE_VALUE) {
            Sleep(250);
            continue;
        }

        const BOOL connected = ConnectNamedPipe(pipe, nullptr) ? TRUE : (GetLastError() == ERROR_PIPE_CONNECTED);
        if (!connected) {
            CloseHandle(pipe);
            Sleep(100);
            continue;
        }

        std::string cmd;
        while (running.load(std::memory_order_acquire) && ReadCommand(pipe, cmd)) {
            if (_stricmp(cmd.c_str(), "PING") == 0) {
                WriteResponse(pipe, "PONG\n");
                continue;
            }
            if (_stricmp(cmd.c_str(), "QUIT") == 0) {
                WriteResponse(pipe, "BYE\n");
                running.store(false, std::memory_order_release);
                break;
            }
            // Current host release is intentionally conservative:
            // command channel is stable, VST processing integration follows
            // through controlled protocol/version upgrades.
            WriteResponse(pipe, "OK\n");
        }

        FlushFileBuffers(pipe);
        DisconnectNamedPipe(pipe);
        CloseHandle(pipe);
    }
    return 0;
}
