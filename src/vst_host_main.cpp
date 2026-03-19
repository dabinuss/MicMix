#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <objbase.h>
#include <ole2.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <condition_variable>
#include <cstdint>
#include <cstring>
#include <deque>
#include <filesystem>
#include <fstream>
#include <memory>
#include <mutex>
#include <sstream>
#include <string>
#include <thread>
#include <unordered_map>
#include <vector>
#include <xmmintrin.h>
#include <pmmintrin.h>

#include "vst_audio_ipc.h"

#define INIT_CLASS_IID 1
#include "pluginterfaces/base/funknown.h"
#include "pluginterfaces/base/ipluginbase.h"
#include "pluginterfaces/gui/iplugview.h"
#include "pluginterfaces/vst/ivstaudioprocessor.h"
#include "pluginterfaces/vst/ivstcomponent.h"
#include "pluginterfaces/vst/ivsteditcontroller.h"
#include "pluginterfaces/vst/ivsthostapplication.h"
#include "pluginterfaces/vst/ivstmessage.h"
#undef INIT_CLASS_IID

namespace {

using Steinberg::FUnknown;
using Steinberg::IPluginFactory;
using Steinberg::IPluginFactory2;
using Steinberg::PClassInfo;
using Steinberg::PClassInfo2;
using Steinberg::TUID;
using Steinberg::ViewRect;
using Steinberg::kPlatformTypeHWND;
using Steinberg::kResultOk;
using Steinberg::kResultTrue;
using Steinberg::tresult;
using Steinberg::Vst::AudioBusBuffers;
using Steinberg::Vst::BusInfo;
using Steinberg::Vst::IAudioProcessor;
using Steinberg::Vst::IComponent;
using Steinberg::Vst::IComponentHandler;
using Steinberg::Vst::IConnectionPoint;
using Steinberg::Vst::IEditController;
using Steinberg::Vst::IHostApplication;
using Steinberg::Vst::ProcessData;
using Steinberg::Vst::ProcessSetup;
using Steinberg::Vst::SpeakerArrangement;
using Steinberg::Vst::kAudio;
using Steinberg::Vst::kInput;
using Steinberg::Vst::kOutput;
using Steinberg::Vst::kRealtime;
using Steinberg::Vst::kSample32;

constexpr size_t kMaxEffectsPerChain = 64;
constexpr size_t kMaxPipeMessageBytes = 64U * 1024U;
constexpr uint32_t kProcessTimeoutMs = 30U;
constexpr UINT kMsgUiWake = WM_APP + 201U;

struct EffectSlot {
    std::string path;
    std::string name;
    std::string uid;
    std::string stateBlob;
    std::string lastStatus;
    bool enabled = true;
    bool bypass = false;
};

std::wstring Utf8ToWide(const std::string& text) {
    if (text.empty()) {
        return {};
    }
    const int len = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, nullptr, 0);
    if (len <= 0) {
        return {};
    }
    std::wstring out(static_cast<size_t>(len), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, out.data(), len);
    if (!out.empty() && out.back() == L'\0') {
        out.pop_back();
    }
    return out;
}

std::string Trim(std::string v) {
    while (!v.empty() && (v.back() == '\r' || v.back() == '\n' || v.back() == '\t' || v.back() == ' ')) {
        v.pop_back();
    }
    size_t i = 0;
    while (i < v.size() && (v[i] == '\r' || v[i] == '\n' || v[i] == '\t' || v[i] == ' ')) {
        ++i;
    }
    return v.substr(i);
}

std::string ToToken(std::string value) {
    for (char& ch : value) {
        if (ch <= ' ' || ch == '=' || ch == '\r' || ch == '\n' || ch == '\t') {
            ch = '_';
        }
    }
    if (value.empty()) {
        return "unknown";
    }
    return value;
}

std::filesystem::path BuildHostLogPath() {
    wchar_t appData[MAX_PATH]{};
    const DWORD len = GetEnvironmentVariableW(L"APPDATA", appData, static_cast<DWORD>(std::size(appData)));
    if (len == 0 || len >= std::size(appData)) {
        return std::filesystem::path(L"vst_host.log");
    }
    std::filesystem::path dir = std::filesystem::path(appData) / L"TS3Client" / L"plugins" / L"micmix";
    std::error_code ec;
    std::filesystem::create_directories(dir, ec);
    return dir / L"vst_host.log";
}

void HostLog(const std::string& text) {
    static const std::filesystem::path logPath = BuildHostLogPath();
    static std::mutex logMutex;
    SYSTEMTIME st{};
    GetLocalTime(&st);
    std::lock_guard<std::mutex> lock(logMutex);
    std::ofstream out(logPath, std::ios::app);
    if (!out.is_open()) {
        return;
    }
    out << st.wYear << '-'
        << static_cast<int>(st.wMonth) << '-'
        << static_cast<int>(st.wDay) << ' '
        << static_cast<int>(st.wHour) << ':'
        << static_cast<int>(st.wMinute) << ':'
        << static_cast<int>(st.wSecond) << '.'
        << static_cast<int>(st.wMilliseconds)
        << " [HOST] " << text << "\n";
}

bool ParseBool(const std::string& value, bool fallback) {
    if (_stricmp(value.c_str(), "1") == 0 || _stricmp(value.c_str(), "true") == 0) {
        return true;
    }
    if (_stricmp(value.c_str(), "0") == 0 || _stricmp(value.c_str(), "false") == 0) {
        return false;
    }
    return fallback;
}

int HexNibble(char c) {
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return 10 + c - 'a';
    if (c >= 'A' && c <= 'F') return 10 + c - 'A';
    return -1;
}

bool HexDecode(const std::string& hex, std::string& out) {
    out.clear();
    if ((hex.size() % 2U) != 0U) {
        return false;
    }
    out.reserve(hex.size() / 2U);
    for (size_t i = 0; i < hex.size(); i += 2U) {
        const int hi = HexNibble(hex[i]);
        const int lo = HexNibble(hex[i + 1U]);
        if (hi < 0 || lo < 0) {
            out.clear();
            return false;
        }
        out.push_back(static_cast<char>((hi << 4) | lo));
    }
    return true;
}

std::string HexEncode(const std::string& in) {
    static constexpr char kHex[] = "0123456789ABCDEF";
    std::string out;
    out.reserve(in.size() * 2U);
    for (unsigned char ch : in) {
        out.push_back(kHex[(ch >> 4) & 0x0F]);
        out.push_back(kHex[ch & 0x0F]);
    }
    return out;
}

bool ValidateVstPath(const std::string& path, std::string& error) {
    error.clear();
    if (path.empty()) {
        error = "empty_path";
        return false;
    }
    std::filesystem::path p(Utf8ToWide(path));
    if (p.empty() || !p.is_absolute()) {
        error = "path_not_absolute";
        return false;
    }
    const std::wstring raw = p.wstring();
    if (raw.rfind(L"\\\\", 0) == 0) {
        error = "network_path_blocked";
        return false;
    }
    std::error_code ec;
    const std::filesystem::path canonical = std::filesystem::weakly_canonical(p, ec);
    const std::filesystem::path resolved = ec ? p.lexically_normal() : canonical;
    if (resolved.empty()) {
        error = "invalid_path";
        return false;
    }
    if (_wcsicmp(resolved.extension().c_str(), L".vst3") != 0) {
        error = "unsupported_extension";
        return false;
    }
    if (!std::filesystem::exists(resolved, ec) || ec || std::filesystem::is_directory(resolved, ec)) {
        error = "path_missing";
        return false;
    }
    return true;
}

struct HostApplication final : IHostApplication {
    tresult PLUGIN_API queryInterface(const TUID iidValue, void** obj) override {
        if (!obj) {
            return Steinberg::kInvalidArgument;
        }
        *obj = nullptr;
        if (Steinberg::FUnknownPrivate::iidEqual(iidValue, IHostApplication::iid) ||
            Steinberg::FUnknownPrivate::iidEqual(iidValue, FUnknown::iid)) {
            *obj = static_cast<IHostApplication*>(this);
            addRef();
            return kResultOk;
        }
        return Steinberg::kNoInterface;
    }

    // Host-owned lifetime: plug-ins must not delete host context objects.
    Steinberg::uint32 PLUGIN_API addRef() override { return 1000; }
    Steinberg::uint32 PLUGIN_API release() override { return 1000; }

    tresult PLUGIN_API getName(Steinberg::Vst::String128 name) override {
        std::fill(name, name + 128, 0);
        static constexpr const char* kHostName = "MicMix Host";
        for (size_t i = 0; kHostName[i] != '\0' && i < 127U; ++i) {
            name[i] = static_cast<Steinberg::char16>(kHostName[i]);
        }
        return kResultOk;
    }

    tresult PLUGIN_API createInstance(TUID, TUID, void**) override {
        return Steinberg::kNoInterface;
    }
};

class HostComponentHandler final : public IComponentHandler {
public:
    tresult PLUGIN_API queryInterface(const TUID iidValue, void** obj) override {
        if (!obj) {
            return Steinberg::kInvalidArgument;
        }
        *obj = nullptr;
        if (Steinberg::FUnknownPrivate::iidEqual(iidValue, IComponentHandler::iid) ||
            Steinberg::FUnknownPrivate::iidEqual(iidValue, FUnknown::iid)) {
            *obj = static_cast<IComponentHandler*>(this);
            addRef();
            return kResultOk;
        }
        return Steinberg::kNoInterface;
    }

    // Host-owned lifetime: plug-ins must not delete host callback objects.
    Steinberg::uint32 PLUGIN_API addRef() override { return 1000; }
    Steinberg::uint32 PLUGIN_API release() override { return 1000; }

    tresult PLUGIN_API beginEdit(Steinberg::Vst::ParamID) override {
        return kResultOk;
    }

    tresult PLUGIN_API performEdit(Steinberg::Vst::ParamID, Steinberg::Vst::ParamValue) override {
        return kResultOk;
    }

    tresult PLUGIN_API endEdit(Steinberg::Vst::ParamID) override {
        return kResultOk;
    }

    tresult PLUGIN_API restartComponent(Steinberg::int32) override {
        return kResultOk;
    }
};

class LoadedVst {
public:
    ~LoadedVst() {
        Shutdown();
    }

    bool Load(const std::string& path, HostApplication* hostContext, std::string& error);
    bool Process(float* mono, uint32_t frames);
    bool OpenEditor(const std::wstring& title, std::string& error);
    void CloseEditor();
    void Shutdown();

private:
    class EditorFrame;
    static LRESULT CALLBACK EditorWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static const wchar_t* EditorWindowClassName();
    void NotifyEditorInit(bool ok, const std::string& errorText);
    bool OnEditorCreateWindow(HWND hwnd);
    void OnEditorInit(HWND hwnd);
    void OnEditorSize();
    void OnEditorDestroy();
    void TryLoadController();

    HMODULE module_ = nullptr;
    IPluginFactory* factory_ = nullptr;
    IComponent* component_ = nullptr;
    IAudioProcessor* processor_ = nullptr;
    IEditController* controller_ = nullptr;
    IEditController* componentController_ = nullptr;
    IComponentHandler* componentHandler_ = nullptr;
    IConnectionPoint* componentConnection_ = nullptr;
    IConnectionPoint* controllerConnection_ = nullptr;
    bool (*exitDll_)() = nullptr;
    HostApplication* hostContext_ = nullptr;
    TUID controllerClassId_{};
    bool controllerClassIdValid_ = false;

    Steinberg::int32 inChannels_ = 1;
    Steinberg::int32 outChannels_ = 1;
    std::vector<float> inA_;
    std::vector<float> inB_;
    std::vector<float> outA_;
    std::vector<float> outB_;

    HWND editorHwnd_ = nullptr;
    std::wstring editorTitle_;
    bool editorInitOk_ = false;
    std::string editorInitError_;
    Steinberg::IPlugView* plugView_ = nullptr;
    EditorFrame* plugFrame_ = nullptr;
};

class LoadedVst::EditorFrame final : public Steinberg::IPlugFrame {
public:
    explicit EditorFrame(HWND hwnd) : hwnd_(hwnd) {}

    tresult PLUGIN_API queryInterface(const TUID iidValue, void** obj) override {
        if (!obj) {
            return Steinberg::kInvalidArgument;
        }
        *obj = nullptr;
        if (Steinberg::FUnknownPrivate::iidEqual(iidValue, Steinberg::IPlugFrame::iid) ||
            Steinberg::FUnknownPrivate::iidEqual(iidValue, FUnknown::iid)) {
            *obj = static_cast<Steinberg::IPlugFrame*>(this);
            addRef();
            return kResultOk;
        }
        return Steinberg::kNoInterface;
    }

    // Host-owned lifetime: plug-ins must not delete host frame objects.
    Steinberg::uint32 PLUGIN_API addRef() override { return 1000; }
    Steinberg::uint32 PLUGIN_API release() override { return 1000; }

    tresult PLUGIN_API resizeView(Steinberg::IPlugView* view, ViewRect* newSize) override {
        if (!hwnd_ || !view || !newSize) {
            return Steinberg::kInvalidArgument;
        }
        if (resizeRecursionGuard_) {
            return Steinberg::kResultFalse;
        }
        ViewRect currentRect{};
        if (view->getSize(&currentRect) == kResultOk) {
            if (currentRect.left == newSize->left &&
                currentRect.top == newSize->top &&
                currentRect.right == newSize->right &&
                currentRect.bottom == newSize->bottom) {
                return kResultOk;
            }
        }

        resizeRecursionGuard_ = true;
        RECT wr{0, 0, newSize->getWidth(), newSize->getHeight()};
        AdjustWindowRectEx(&wr, WS_OVERLAPPEDWINDOW, FALSE, 0);
        SetWindowPos(
            hwnd_,
            nullptr,
            0,
            0,
            wr.right - wr.left,
            wr.bottom - wr.top,
            SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        resizeRecursionGuard_ = false;

        if (view->getSize(&currentRect) == kResultOk) {
            if (currentRect.left != newSize->left ||
                currentRect.top != newSize->top ||
                currentRect.right != newSize->right ||
                currentRect.bottom != newSize->bottom) {
                view->onSize(newSize);
            }
        }
        return kResultOk;
    }

private:
    HWND hwnd_ = nullptr;
    bool resizeRecursionGuard_ = false;
};

bool LoadedVst::Load(const std::string& path, HostApplication* hostContext, std::string& error) {
    using InitDllProc = bool (*)();
    using ExitDllProc = bool (*)();
    using GetFactoryProc = IPluginFactory* (*)();

    Shutdown();
    error.clear();
    if (!hostContext) {
        error = "missing_host_context";
        return false;
    }
    hostContext_ = hostContext;
    controllerClassIdValid_ = false;
    std::memset(controllerClassId_, 0, sizeof(controllerClassId_));
    if (!ValidateVstPath(path, error)) {
        return false;
    }

    module_ = LoadLibraryW(Utf8ToWide(path).c_str());
    if (!module_) {
        error = "LoadLibrary_failed";
        return false;
    }

    if (auto* initDll = reinterpret_cast<InitDllProc>(GetProcAddress(module_, "InitDll"))) {
        initDll();
    }
    exitDll_ = reinterpret_cast<ExitDllProc>(GetProcAddress(module_, "ExitDll"));

    auto* getFactory = reinterpret_cast<GetFactoryProc>(GetProcAddress(module_, "GetPluginFactory"));
    if (!getFactory) {
        error = "GetPluginFactory_missing";
        Shutdown();
        return false;
    }

    factory_ = getFactory();
    if (!factory_) {
        error = "factory_null";
        Shutdown();
        return false;
    }

    TUID classId{};
    bool foundEffectClass = false;
    Steinberg::FUnknownPtr<IPluginFactory2> factory2(factory_);
    const Steinberg::int32 classCount = factory_->countClasses();
    for (Steinberg::int32 i = 0; i < classCount; ++i) {
        if (factory2) {
            PClassInfo2 info2{};
            if (factory2->getClassInfo2(i, &info2) == kResultOk) {
                if (std::strcmp(info2.category, kVstAudioEffectClass) == 0) {
                    std::memcpy(classId, info2.cid, sizeof(TUID));
                    foundEffectClass = true;
                    break;
                }
            }
        }
        PClassInfo info1{};
        if (factory_->getClassInfo(i, &info1) == kResultOk) {
            if (std::strcmp(info1.category, kVstAudioEffectClass) == 0) {
                std::memcpy(classId, info1.cid, sizeof(TUID));
                foundEffectClass = true;
                break;
            }
        }
    }

    if (!foundEffectClass) {
        error = "audio_effect_class_missing";
        Shutdown();
        return false;
    }

    TUID componentIid{};
    IComponent::iid.toTUID(componentIid);
    if (factory_->createInstance(classId, componentIid, reinterpret_cast<void**>(&component_)) != kResultOk || !component_) {
        error = "component_create_failed";
        Shutdown();
        return false;
    }

    if (component_->initialize(hostContext) != kResultOk) {
        error = "component_initialize_failed";
        Shutdown();
        return false;
    }

    TUID processorIid{};
    IAudioProcessor::iid.toTUID(processorIid);
    if (component_->queryInterface(processorIid, reinterpret_cast<void**>(&processor_)) != kResultOk || !processor_) {
        error = "audio_processor_missing";
        Shutdown();
        return false;
    }

    if (component_->getControllerClassId(controllerClassId_) == kResultOk) {
        controllerClassIdValid_ = true;
    }

    BusInfo inBusInfo{};
    BusInfo outBusInfo{};
    if (component_->getBusInfo(kAudio, kInput, 0, inBusInfo) != kResultOk ||
        component_->getBusInfo(kAudio, kOutput, 0, outBusInfo) != kResultOk) {
        error = "bus_info_failed";
        Shutdown();
        return false;
    }

    inChannels_ = std::clamp<Steinberg::int32>(inBusInfo.channelCount, 1, 2);
    outChannels_ = std::clamp<Steinberg::int32>(outBusInfo.channelCount, 1, 2);

    SpeakerArrangement inArr[1] = {
        (inChannels_ > 1) ? Steinberg::Vst::SpeakerArr::kStereo : Steinberg::Vst::SpeakerArr::kMono};
    SpeakerArrangement outArr[1] = {
        (outChannels_ > 1) ? Steinberg::Vst::SpeakerArr::kStereo : Steinberg::Vst::SpeakerArr::kMono};
    if (processor_->setBusArrangements(inArr, 1, outArr, 1) != kResultOk) {
        error = "bus_arrangement_failed";
        Shutdown();
        return false;
    }

    if (component_->activateBus(kAudio, kInput, 0, 1) != kResultOk ||
        component_->activateBus(kAudio, kOutput, 0, 1) != kResultOk) {
        error = "activate_bus_failed";
        Shutdown();
        return false;
    }

    if (processor_->canProcessSampleSize(kSample32) != kResultTrue) {
        error = "float32_not_supported";
        Shutdown();
        return false;
    }

    ProcessSetup setup{};
    setup.processMode = kRealtime;
    setup.symbolicSampleSize = kSample32;
    setup.maxSamplesPerBlock = static_cast<Steinberg::int32>(micmix::vstipc::kMaxFramesPerPacket);
    setup.sampleRate = 48000.0;
    if (processor_->setupProcessing(setup) != kResultOk) {
        error = "setup_processing_failed";
        Shutdown();
        return false;
    }

    if (component_->setActive(1) != kResultOk) {
        error = "set_active_failed";
        Shutdown();
        return false;
    }

    if (processor_->setProcessing(1) != kResultOk) {
        error = "set_processing_failed";
        Shutdown();
        return false;
    }

    inA_.assign(micmix::vstipc::kMaxFramesPerPacket, 0.0f);
    inB_.assign(micmix::vstipc::kMaxFramesPerPacket, 0.0f);
    outA_.assign(micmix::vstipc::kMaxFramesPerPacket, 0.0f);
    outB_.assign(micmix::vstipc::kMaxFramesPerPacket, 0.0f);
    return true;
}

bool LoadedVst::Process(float* mono, uint32_t frames) {
    if (!processor_ || !mono || frames == 0 || frames > micmix::vstipc::kMaxFramesPerPacket) {
        return false;
    }

    for (uint32_t i = 0; i < frames; ++i) {
        const float in = mono[i];
        inA_[i] = in;
        inB_[i] = in;
        outA_[i] = 0.0f;
        outB_[i] = 0.0f;
    }

    float* inPtrs[2] = {inA_.data(), inB_.data()};
    float* outPtrs[2] = {outA_.data(), outB_.data()};

    AudioBusBuffers inBus{};
    inBus.numChannels = inChannels_;
    inBus.silenceFlags = 0;
    inBus.channelBuffers32 = inPtrs;

    AudioBusBuffers outBus{};
    outBus.numChannels = outChannels_;
    outBus.silenceFlags = 0;
    outBus.channelBuffers32 = outPtrs;

    ProcessData processData{};
    processData.processMode = kRealtime;
    processData.symbolicSampleSize = kSample32;
    processData.numSamples = static_cast<Steinberg::int32>(frames);
    processData.numInputs = 1;
    processData.numOutputs = 1;
    processData.inputs = &inBus;
    processData.outputs = &outBus;

    if (processor_->process(processData) != kResultOk) {
        return false;
    }

    if (outChannels_ <= 1) {
        for (uint32_t i = 0; i < frames; ++i) {
            float v = outA_[i];
            if (!std::isfinite(v)) {
                v = 0.0f;
            }
            mono[i] = std::clamp(v, -8.0f, 8.0f);
        }
    } else {
        float sumAbsL = 0.0f;
        float sumAbsR = 0.0f;
        float sumAbsMid = 0.0f;
        float sumAbsSide = 0.0f;
        for (uint32_t i = 0; i < frames; ++i) {
            const float l = std::isfinite(outA_[i]) ? outA_[i] : 0.0f;
            const float r = std::isfinite(outB_[i]) ? outB_[i] : 0.0f;
            sumAbsL += std::fabs(l);
            sumAbsR += std::fabs(r);
            sumAbsMid += std::fabs(l + r);
            sumAbsSide += std::fabs(l - r);
        }
        const bool preferDominantChannel = (sumAbsSide > (sumAbsMid * 3.0f));
        const bool chooseLeft = (sumAbsL >= sumAbsR);
        for (uint32_t i = 0; i < frames; ++i) {
            const float l = std::isfinite(outA_[i]) ? outA_[i] : 0.0f;
            const float r = std::isfinite(outB_[i]) ? outB_[i] : 0.0f;
            float downmix = preferDominantChannel ? (chooseLeft ? l : r) : (0.5f * (l + r));
            if (!std::isfinite(downmix)) {
                downmix = 0.0f;
            }
            mono[i] = std::clamp(downmix, -8.0f, 8.0f);
        }
    }
    return true;
}

void LoadedVst::TryLoadController() {
    if (controller_ || componentController_) {
        return;
    }
    if (!component_ || !hostContext_) {
        return;
    }
    if (controllerClassIdValid_ && factory_) {
        TUID controllerIid{};
        IEditController::iid.toTUID(controllerIid);
        IEditController* controller = nullptr;
        if (factory_->createInstance(controllerClassId_, controllerIid, reinterpret_cast<void**>(&controller)) == kResultOk &&
            controller) {
            if (controller->initialize(hostContext_) == kResultOk) {
                controller_ = controller;
            } else {
                controller->release();
            }
        }
    }
    if (!controller_) {
        TUID controllerIid{};
        IEditController::iid.toTUID(controllerIid);
        IEditController* componentController = nullptr;
        if (component_->queryInterface(controllerIid, reinterpret_cast<void**>(&componentController)) == kResultOk &&
            componentController) {
            componentController_ = componentController;
        }
    }
    if (!componentHandler_) {
        componentHandler_ = reinterpret_cast<IComponentHandler*>(new HostComponentHandler());
    }
    if (componentHandler_) {
        if (controller_) {
            controller_->setComponentHandler(componentHandler_);
        }
        if (componentController_) {
            componentController_->setComponentHandler(componentHandler_);
        }
    }
    if (!componentConnection_ && !controllerConnection_ && controller_) {
        TUID connectionIid{};
        IConnectionPoint::iid.toTUID(connectionIid);
        IConnectionPoint* compCp = nullptr;
        IConnectionPoint* ctrlCp = nullptr;
        if (component_->queryInterface(connectionIid, reinterpret_cast<void**>(&compCp)) == kResultOk && compCp &&
            controller_->queryInterface(connectionIid, reinterpret_cast<void**>(&ctrlCp)) == kResultOk && ctrlCp) {
            compCp->connect(ctrlCp);
            ctrlCp->connect(compCp);
            componentConnection_ = compCp;
            controllerConnection_ = ctrlCp;
        } else {
            if (compCp) {
                compCp->release();
            }
            if (ctrlCp) {
                ctrlCp->release();
            }
        }
    }
}

LRESULT CALLBACK LoadedVst::EditorWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    LoadedVst* self = reinterpret_cast<LoadedVst*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE) {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
        self = reinterpret_cast<LoadedVst*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }
    if (!self) {
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
    case WM_CREATE:
        return self->OnEditorCreateWindow(hwnd) ? 0 : -1;
    case WM_SIZE:
        self->OnEditorSize();
        return 0;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
        self->OnEditorDestroy();
        return 0;
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
}

const wchar_t* LoadedVst::EditorWindowClassName() {
    static const wchar_t* kClassName = L"MicMixVstEditorHostWindow";
    static std::once_flag once;
    std::call_once(once, []() {
        WNDCLASSW wc{};
        wc.lpfnWndProc = &LoadedVst::EditorWndProc;
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kClassName;
        RegisterClassW(&wc);
    });
    return kClassName;
}

void LoadedVst::NotifyEditorInit(bool ok, const std::string& errorText) {
    editorInitOk_ = ok;
    editorInitError_ = errorText;
    HostLog(std::string("editor_init ") + (ok ? "ok" : "fail") +
            (errorText.empty() ? std::string() : (" reason=" + ToToken(errorText))));
}

bool LoadedVst::OnEditorCreateWindow(HWND hwnd) {
    editorHwnd_ = hwnd;
    return true;
}

void LoadedVst::OnEditorInit(HWND hwnd) {
    HostLog("editor_init_step begin");
    TryLoadController();
    if (!controller_ && !componentController_) {
        NotifyEditorInit(false, "editor_not_supported");
        DestroyWindow(hwnd);
        return;
    }

    HostLog("editor_init_step create_view");
    Steinberg::IPlugView* view = nullptr;
    if (controller_) {
        view = controller_->createView(Steinberg::Vst::ViewType::kEditor);
    }
    if (!view && componentController_) {
        view = componentController_->createView(Steinberg::Vst::ViewType::kEditor);
    }
    if (!view) {
        NotifyEditorInit(false, "editor_view_missing");
        DestroyWindow(hwnd);
        return;
    }
    HostLog("editor_init_step view_created");
    if (view->isPlatformTypeSupported(kPlatformTypeHWND) != kResultTrue) {
        view->release();
        NotifyEditorInit(false, "platform_hwnd_not_supported");
        DestroyWindow(hwnd);
        return;
    }

    auto* frame = new EditorFrame(hwnd);
    if (view->setFrame(frame) != kResultOk) {
        delete frame;
        view->release();
        NotifyEditorInit(false, "set_frame_failed");
        DestroyWindow(hwnd);
        return;
    }

    const HWND platformParent = hwnd;

    ViewRect rect{};
    if (view->getSize(&rect) != kResultOk || rect.getWidth() <= 0 || rect.getHeight() <= 0) {
        rect = ViewRect(0, 0, 720, 520);
    }
    RECT wr{0, 0, rect.getWidth(), rect.getHeight()};
    AdjustWindowRectEx(&wr, WS_OVERLAPPEDWINDOW, FALSE, 0);
    SetWindowPos(
        hwnd,
        nullptr,
        0,
        0,
        wr.right - wr.left,
        wr.bottom - wr.top,
        SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    HostLog("editor_init_step attach");
    if (view->attached(platformParent, kPlatformTypeHWND) != kResultOk) {
        view->setFrame(nullptr);
        delete frame;
        view->release();
        NotifyEditorInit(false, "view_attach_failed");
        DestroyWindow(hwnd);
        return;
    }

    plugView_ = view;
    plugFrame_ = frame;
    NotifyEditorInit(true, {});
    HostLog("editor_init_step done");
}

void LoadedVst::OnEditorSize() {
    Steinberg::IPlugView* view = plugView_;
    if (!view) {
        return;
    }
    HWND hwnd = editorHwnd_;
    if (!hwnd) {
        return;
    }
    RECT client{};
    GetClientRect(hwnd, &client);
    ViewRect rect(0, 0, client.right - client.left, client.bottom - client.top);
    ViewRect current{};
    if (view->getSize(&current) == kResultOk) {
        if (current.left == rect.left &&
            current.top == rect.top &&
            current.right == rect.right &&
            current.bottom == rect.bottom) {
            return;
        }
    }
    view->onSize(&rect);
}

void LoadedVst::OnEditorDestroy() {
    Steinberg::IPlugView* view = plugView_;
    EditorFrame* frame = plugFrame_;
    plugView_ = nullptr;
    plugFrame_ = nullptr;
    if (view) {
        view->setFrame(nullptr);
        view->removed();
        view->release();
    }
    if (frame) {
        delete frame;
    }
    editorHwnd_ = nullptr;
}

bool LoadedVst::OpenEditor(const std::wstring& title, std::string& error) {
    error.clear();
    const HWND existing = editorHwnd_;
    if (existing) {
        ShowWindow(existing, SW_RESTORE);
        SetForegroundWindow(existing);
        return true;
    }
    editorTitle_ = title.empty() ? L"MicMix VST Editor" : title;
    editorInitOk_ = false;
    editorInitError_.clear();

    HWND hwnd = CreateWindowExW(
        WS_EX_APPWINDOW,
        EditorWindowClassName(),
        editorTitle_.c_str(),
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        860,
        620,
        nullptr,
        nullptr,
        GetModuleHandleW(nullptr),
        this);
    if (!hwnd) {
        error = "editor_window_create_failed";
        HostLog("editor_window_create_failed gle=" + std::to_string(GetLastError()));
        return false;
    }
    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);
    OnEditorInit(hwnd);
    if (!editorInitOk_) {
        error = editorInitError_.empty() ? "editor_open_failed" : editorInitError_;
        if (editorHwnd_ && IsWindow(editorHwnd_)) {
            DestroyWindow(editorHwnd_);
        }
        return false;
    }
    return true;
}

void LoadedVst::CloseEditor() {
    const HWND hwnd = editorHwnd_;
    if (hwnd && IsWindow(hwnd)) {
        DestroyWindow(hwnd);
    }
    editorHwnd_ = nullptr;
}

void LoadedVst::Shutdown() {
    CloseEditor();

    if (componentConnection_ && controllerConnection_) {
        controllerConnection_->disconnect(componentConnection_);
        componentConnection_->disconnect(controllerConnection_);
    }
    if (controllerConnection_) {
        controllerConnection_->release();
        controllerConnection_ = nullptr;
    }
    if (componentConnection_) {
        componentConnection_->release();
        componentConnection_ = nullptr;
    }
    if (componentHandler_) {
        if (controller_) {
            controller_->setComponentHandler(nullptr);
        }
        if (componentController_) {
            componentController_->setComponentHandler(nullptr);
        }
        delete static_cast<HostComponentHandler*>(componentHandler_);
        componentHandler_ = nullptr;
    }

    if (componentController_) {
        componentController_->release();
        componentController_ = nullptr;
    }
    if (controller_) {
        controller_->terminate();
        controller_->release();
        controller_ = nullptr;
    }

    if (processor_) {
        processor_->setProcessing(0);
    }
    if (component_) {
        component_->setActive(0);
        component_->terminate();
    }
    if (processor_) {
        processor_->release();
        processor_ = nullptr;
    }
    if (component_) {
        component_->release();
        component_ = nullptr;
    }
    if (factory_) {
        factory_->release();
        factory_ = nullptr;
    }
    if (module_) {
        if (exitDll_) {
            exitDll_();
        }
        FreeLibrary(module_);
        module_ = nullptr;
    }
    exitDll_ = nullptr;
    hostContext_ = nullptr;
    controllerClassIdValid_ = false;
    std::memset(controllerClassId_, 0, sizeof(controllerClassId_));

    inChannels_ = 1;
    outChannels_ = 1;
    inA_.clear();
    inB_.clear();
    outA_.clear();
    outB_.clear();
}

enum class UiCommandType {
    Sync,
    OpenEditor,
    CloseEditor,
    Quit,
};

struct UiSyncRequest {
    bool effectsEnabled = false;
    std::vector<EffectSlot> musicSlots;
    std::vector<EffectSlot> micSlots;
    std::mutex mutex;
    std::condition_variable cv;
    bool done = false;
    std::string response;
};

struct UiCommand {
    UiCommandType type = UiCommandType::Sync;
    bool isMusic = false;
    size_t index = 0;
    std::shared_ptr<UiSyncRequest> sync;
};

struct HostState {
    bool effectsEnabled = false;
    std::vector<EffectSlot> musicSlots;
    std::vector<EffectSlot> micSlots;
    std::vector<int> musicSlotToLoaded;
    std::vector<int> micSlotToLoaded;
    std::vector<std::unique_ptr<LoadedVst>> musicChain;
    std::vector<std::unique_ptr<LoadedVst>> micChain;
    std::mutex chainMutex;
    std::string status = "idle";
    uint32_t blocked = 0;

    HANDLE shmMap = nullptr;
    micmix::vstipc::SharedMemory* shm = nullptr;
    std::atomic<bool> workerRun{false};
    std::thread worker;

    std::mutex uiCmdMutex;
    std::deque<UiCommand> uiCmdQueue;
    std::atomic<bool> uiCmdRun{false};
    std::thread uiCmdThread;
    std::atomic<DWORD> uiThreadId{0};

    HostApplication* hostContext = nullptr;
};

bool ReadMessage(HANDLE pipe, std::string& outLine) {
    outLine.clear();
    std::array<char, 4096> buffer{};
    size_t total = 0;
    for (;;) {
        DWORD read = 0;
        const BOOL ok = ReadFile(pipe, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr);
        if (read > 0) {
            total += static_cast<size_t>(read);
            if (total > kMaxPipeMessageBytes) {
                return false;
            }
            outLine.append(buffer.data(), buffer.data() + read);
        }
        if (ok) {
            break;
        }
        if (GetLastError() != ERROR_MORE_DATA) {
            return false;
        }
    }
    outLine = Trim(outLine);
    return !outLine.empty();
}

void WriteResponse(HANDLE pipe, const std::string& text) {
    const std::string wire = text + "\n";
    DWORD written = 0;
    WriteFile(pipe, wire.data(), static_cast<DWORD>(wire.size()), &written, nullptr);
}

void ResetAudioRings(HostState& state) {
    if (!state.shm) {
        return;
    }
    micmix::vstipc::RingReset(state.shm->musicIn);
    micmix::vstipc::RingReset(state.shm->musicOut);
    micmix::vstipc::RingReset(state.shm->micIn);
    micmix::vstipc::RingReset(state.shm->micOut);
    InterlockedExchange(&state.shm->hostHeartbeat, static_cast<LONG>(GetTickCount() & 0x7FFFFFFF));
    InterlockedExchange(&state.shm->pluginHeartbeat, 0);
}

void RebuildChains(HostState& state) {
    std::lock_guard<std::mutex> lock(state.chainMutex);
    state.musicChain.clear();
    state.micChain.clear();
    state.musicSlotToLoaded.assign(state.musicSlots.size(), -1);
    state.micSlotToLoaded.assign(state.micSlots.size(), -1);
    state.blocked = 0;
    state.status = "running";

    auto loadOne = [&](const EffectSlot& slot, size_t slotIndex, std::vector<int>& slotMap, std::vector<std::unique_ptr<LoadedVst>>& chain) {
        if (!slot.enabled || slot.bypass) {
            return;
        }
        auto plugin = std::make_unique<LoadedVst>();
        std::string error;
        if (plugin->Load(slot.path, state.hostContext, error)) {
            slotMap[slotIndex] = static_cast<int>(chain.size());
            chain.push_back(std::move(plugin));
            return;
        }
        ++state.blocked;
        state.status = "degraded_bypass";
    };

    for (size_t i = 0; i < state.musicSlots.size(); ++i) {
        loadOne(state.musicSlots[i], i, state.musicSlotToLoaded, state.musicChain);
    }
    for (size_t i = 0; i < state.micSlots.size(); ++i) {
        loadOne(state.micSlots[i], i, state.micSlotToLoaded, state.micChain);
    }

    ResetAudioRings(state);
}

bool ProcessPacketChain(HostState& state, bool isMusic, micmix::vstipc::AudioPacket& packet) {
    packet.frames = std::min<uint32_t>(packet.frames, micmix::vstipc::kMaxFramesPerPacket);
    if (packet.frames == 0) {
        return true;
    }

    std::lock_guard<std::mutex> lock(state.chainMutex);
    if (!state.effectsEnabled) {
        return true;
    }
    const auto& chain = isMusic ? state.musicChain : state.micChain;
    for (const auto& plugin : chain) {
        if (!plugin.get()) {
            continue;
        }
        const auto started = std::chrono::steady_clock::now();
        if (!plugin->Process(packet.samples, packet.frames)) {
            state.status = "degraded_bypass";
            return false;
        }
        const auto elapsedMs = static_cast<uint32_t>(
            std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());
        if (elapsedMs > kProcessTimeoutMs) {
            state.status = "timeout_bypass";
            return false;
        }
    }
    return true;
}

void WorkerMain(HostState* state) {
    if (!state) {
        return;
    }
    _MM_SET_FLUSH_ZERO_MODE(_MM_FLUSH_ZERO_ON);
#if defined(_MM_DENORMALS_ZERO_ON)
    _MM_SET_DENORMALS_ZERO_MODE(_MM_DENORMALS_ZERO_ON);
#endif
    state->workerRun.store(true, std::memory_order_release);
    while (state->workerRun.load(std::memory_order_acquire)) {
        auto* shm = state->shm;
        if (!shm) {
            Sleep(2);
            continue;
        }
        InterlockedExchange(&shm->hostHeartbeat, static_cast<LONG>(GetTickCount() & 0x7FFFFFFF));

        micmix::vstipc::AudioPacket packet{};
        bool didWork = false;
        for (int i = 0; i < 64; ++i) {
            bool progressed = false;
            if (micmix::vstipc::RingPop(shm->micIn, packet)) {
                progressed = true;
                didWork = true;
                ProcessPacketChain(*state, false, packet);
                if (!micmix::vstipc::RingPush(shm->micOut, packet)) {
                    break;
                }
            }
            if (micmix::vstipc::RingPop(shm->musicIn, packet)) {
                progressed = true;
                didWork = true;
                ProcessPacketChain(*state, true, packet);
                if (!micmix::vstipc::RingPush(shm->musicOut, packet)) {
                    break;
                }
            }
            if (!progressed) {
                break;
            }
        }
        if (!didWork) {
            Sleep(1);
        } else {
            SwitchToThread();
        }
    }
}

std::string BuildPong(HostState& state) {
    if (!state.chainMutex.try_lock()) {
        return "PONG busy=1 status=chain_busy";
    }
    std::lock_guard<std::mutex> lock(state.chainMutex, std::adopt_lock);
    std::ostringstream ss;
    ss << "PONG effects=" << (state.effectsEnabled ? 1 : 0) << " music=" << state.musicSlots.size()
       << " mic=" << state.micSlots.size() << " loaded_music=" << state.musicChain.size()
       << " loaded_mic=" << state.micChain.size() << " blocked=" << state.blocked
       << " status=" << state.status;
    return ss.str();
}

void EnqueueUiCommand(HostState& state, UiCommand cmd) {
    {
        std::lock_guard<std::mutex> lock(state.uiCmdMutex);
        state.uiCmdQueue.push_back(std::move(cmd));
    }
    const DWORD uiTid = state.uiThreadId.load(std::memory_order_acquire);
    if (uiTid != 0) {
        PostThreadMessageW(uiTid, kMsgUiWake, 0, 0);
    }
}

std::string HandleSync(const std::string& encodedPayload, HostState& state) {
    std::string payload;
    if (!HexDecode(encodedPayload, payload)) {
        return "ERR invalid_sync_payload_hex";
    }

    std::unordered_map<std::string, std::string> kv;
    std::istringstream lines(payload);
    std::string line;
    while (std::getline(lines, line)) {
        line = Trim(line);
        if (line.empty() || line[0] == '#') {
            continue;
        }
        const size_t eq = line.find('=');
        if (eq == std::string::npos) {
            continue;
        }
        kv[Trim(line.substr(0, eq))] = line.substr(eq + 1);
    }

    auto req = std::make_shared<UiSyncRequest>();
    {
        std::lock_guard<std::mutex> lock(state.chainMutex);
        req->effectsEnabled = state.effectsEnabled;
    }
    if (auto it = kv.find("effects_enabled"); it != kv.end()) {
        req->effectsEnabled = ParseBool(Trim(it->second), req->effectsEnabled);
    }

    auto parseChain = [&](const char* prefix, std::vector<EffectSlot>& out) {
        out.clear();
        size_t count = 0;
        if (auto it = kv.find(std::string(prefix) + ".count"); it != kv.end()) {
            try {
                count = static_cast<size_t>(std::stoull(Trim(it->second)));
            } catch (...) {
                count = 0;
            }
        }
        count = std::min(count, kMaxEffectsPerChain);
        out.reserve(count);
        for (size_t i = 0; i < count; ++i) {
            const std::string base = std::string(prefix) + "." + std::to_string(i);
            auto pathIt = kv.find(base + ".path");
            if (pathIt == kv.end() || Trim(pathIt->second).empty()) {
                continue;
            }
            EffectSlot slot{};
            slot.path = Trim(pathIt->second);
            if (auto it = kv.find(base + ".name"); it != kv.end()) slot.name = Trim(it->second);
            if (auto it = kv.find(base + ".uid"); it != kv.end()) slot.uid = Trim(it->second);
            if (auto it = kv.find(base + ".state_blob"); it != kv.end()) slot.stateBlob = it->second;
            if (auto it = kv.find(base + ".last_status"); it != kv.end()) slot.lastStatus = Trim(it->second);
            if (auto it = kv.find(base + ".enabled"); it != kv.end()) slot.enabled = ParseBool(Trim(it->second), true);
            if (auto it = kv.find(base + ".bypass"); it != kv.end()) slot.bypass = ParseBool(Trim(it->second), false);
            out.push_back(std::move(slot));
        }
    };

    parseChain("music", req->musicSlots);
    parseChain("mic", req->micSlots);

    UiCommand cmd{};
    cmd.type = UiCommandType::Sync;
    cmd.sync = req;
    EnqueueUiCommand(state, std::move(cmd));

    std::unique_lock<std::mutex> lock(req->mutex);
    if (!req->cv.wait_for(lock, std::chrono::milliseconds(5000), [&]() { return req->done; })) {
        return "ERR sync_timeout";
    }
    return req->response.empty() ? "ERR sync_failed" : req->response;
}

std::string HandleScan(const std::string& encodedPath) {
    std::string decodedPath;
    if (!HexDecode(encodedPath, decodedPath) || decodedPath.empty()) {
        return "SCAN_ERR reason=invalid_path";
    }
    std::string error;
    if (!ValidateVstPath(decodedPath, error)) {
        return "SCAN_ERR reason=" + error;
    }
    return "SCAN_OK path=" + HexEncode(decodedPath) + " signed=unknown";
}

bool ParseEditorArgs(const std::string& args, bool& isMusic, size_t& index) {
    std::istringstream ss(args);
    std::string chain;
    size_t parsedIndex = 0;
    if (!(ss >> chain >> parsedIndex)) {
        return false;
    }
    if (_stricmp(chain.c_str(), "music") == 0) {
        isMusic = true;
    } else if (_stricmp(chain.c_str(), "mic") == 0) {
        isMusic = false;
    } else {
        return false;
    }
    index = parsedIndex;
    return true;
}

bool ValidateEditorIndex(HostState& state, bool isMusic, size_t index, std::string& reason) {
    std::lock_guard<std::mutex> lock(state.chainMutex);
    const auto& slots = isMusic ? state.musicSlots : state.micSlots;
    const auto& map = isMusic ? state.musicSlotToLoaded : state.micSlotToLoaded;
    const auto& chain = isMusic ? state.musicChain : state.micChain;
    if (index >= slots.size()) {
        reason = "invalid_index";
        return false;
    }
    if (index >= map.size()) {
        reason = "invalid_slot_map";
        return false;
    }
    const int loadedIndex = map[index];
    if (loadedIndex < 0 || static_cast<size_t>(loadedIndex) >= chain.size() || !chain[static_cast<size_t>(loadedIndex)].get()) {
        reason = "slot_not_loaded";
        return false;
    }
    return true;
}

void ExecuteEditorOpen(HostState& state, bool isMusic, size_t index) {
    LoadedVst* plugin = nullptr;
    std::wstring title;
    {
        std::lock_guard<std::mutex> lock(state.chainMutex);
        const auto& slots = isMusic ? state.musicSlots : state.micSlots;
        const auto& map = isMusic ? state.musicSlotToLoaded : state.micSlotToLoaded;
        const auto& chain = isMusic ? state.musicChain : state.micChain;
        if (index >= slots.size() || index >= map.size()) {
            HostLog("editor_open skip chain=" + std::string(isMusic ? "music" : "mic") + " index=" + std::to_string(index) + " reason=index_changed");
            return;
        }
        const int loadedIndex = map[index];
        if (loadedIndex < 0 || static_cast<size_t>(loadedIndex) >= chain.size() || !chain[static_cast<size_t>(loadedIndex)]) {
            HostLog("editor_open skip chain=" + std::string(isMusic ? "music" : "mic") + " index=" + std::to_string(index) + " reason=slot_not_loaded");
            return;
        }
        plugin = chain[static_cast<size_t>(loadedIndex)].get();
        const auto& slot = slots[index];
        title = L"MicMix VST Editor - " + Utf8ToWide(slot.name.empty() ? slot.path : slot.name);
    }

    std::string error;
    HostLog("editor_open start chain=" + std::string(isMusic ? "music" : "mic") + " index=" + std::to_string(index));
    if (!plugin || !plugin->OpenEditor(title, error)) {
        HostLog("editor_open failed chain=" + std::string(isMusic ? "music" : "mic") + " index=" + std::to_string(index) +
                " reason=" + ToToken(error));
        return;
    }
    HostLog("editor_open ok chain=" + std::string(isMusic ? "music" : "mic") + " index=" + std::to_string(index));
}

void ExecuteEditorClose(HostState& state, bool isMusic, size_t index) {
    LoadedVst* plugin = nullptr;
    {
        std::lock_guard<std::mutex> lock(state.chainMutex);
        const auto& map = isMusic ? state.musicSlotToLoaded : state.micSlotToLoaded;
        const auto& chain = isMusic ? state.musicChain : state.micChain;
        if (index >= map.size()) {
            HostLog("editor_close skip chain=" + std::string(isMusic ? "music" : "mic") + " index=" + std::to_string(index) + " reason=invalid_index");
            return;
        }
        const int loadedIndex = map[index];
        if (loadedIndex < 0 || static_cast<size_t>(loadedIndex) >= chain.size() || !chain[static_cast<size_t>(loadedIndex)]) {
            HostLog("editor_close skip chain=" + std::string(isMusic ? "music" : "mic") + " index=" + std::to_string(index) + " reason=slot_not_loaded");
            return;
        }
        plugin = chain[static_cast<size_t>(loadedIndex)].get();
    }
    if (plugin) {
        plugin->CloseEditor();
    }
}

void UiCommandMain(HostState* state) {
    if (!state) {
        return;
    }
    const HRESULT comHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool comInitOk = SUCCEEDED(comHr);
    const HRESULT oleHr = OleInitialize(nullptr);
    const bool oleInitOk = SUCCEEDED(oleHr);

    state->uiCmdRun.store(true, std::memory_order_release);
    state->uiThreadId.store(GetCurrentThreadId(), std::memory_order_release);
    MSG initMsg{};
    PeekMessageW(&initMsg, nullptr, WM_USER, WM_USER, PM_NOREMOVE);

    bool quit = false;
    while (!quit) {
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) {
                quit = true;
                break;
            }
            if (msg.message == kMsgUiWake) {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        if (quit) {
            break;
        }

        UiCommand cmd{};
        bool hasCommand = false;
        {
            std::lock_guard<std::mutex> lock(state->uiCmdMutex);
            if (!state->uiCmdQueue.empty()) {
                cmd = std::move(state->uiCmdQueue.front());
                state->uiCmdQueue.pop_front();
                hasCommand = true;
            }
        }
        if (hasCommand) {
            switch (cmd.type) {
            case UiCommandType::Sync: {
                if (!cmd.sync) {
                    break;
                }
                {
                    std::lock_guard<std::mutex> lock(state->chainMutex);
                    state->effectsEnabled = cmd.sync->effectsEnabled;
                    state->musicSlots = cmd.sync->musicSlots;
                    state->micSlots = cmd.sync->micSlots;
                }
                RebuildChains(*state);
                std::ostringstream ss;
                {
                    std::lock_guard<std::mutex> lock(state->chainMutex);
                    ss << "SYNC_OK effects=" << (state->effectsEnabled ? 1 : 0)
                       << " music=" << state->musicSlots.size()
                       << " mic=" << state->micSlots.size()
                       << " loaded_music=" << state->musicChain.size()
                       << " loaded_mic=" << state->micChain.size()
                       << " blocked=" << state->blocked
                       << " status=" << state->status;
                }
                {
                    std::lock_guard<std::mutex> doneLock(cmd.sync->mutex);
                    cmd.sync->response = ss.str();
                    cmd.sync->done = true;
                }
                cmd.sync->cv.notify_all();
                break;
            }
            case UiCommandType::OpenEditor:
                ExecuteEditorOpen(*state, cmd.isMusic, cmd.index);
                break;
            case UiCommandType::CloseEditor:
                ExecuteEditorClose(*state, cmd.isMusic, cmd.index);
                break;
            case UiCommandType::Quit:
                {
                    std::lock_guard<std::mutex> lock(state->chainMutex);
                    state->musicChain.clear();
                    state->micChain.clear();
                    state->musicSlotToLoaded.clear();
                    state->micSlotToLoaded.clear();
                    state->status = "idle";
                }
                quit = true;
                break;
            }
            continue;
        }

        if (!state->uiCmdRun.load(std::memory_order_acquire)) {
            bool empty = true;
            {
                std::lock_guard<std::mutex> lock(state->uiCmdMutex);
                empty = state->uiCmdQueue.empty();
            }
            if (empty) {
                break;
            }
        }
        MsgWaitForMultipleObjectsEx(0, nullptr, 30, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
    }

    state->uiThreadId.store(0, std::memory_order_release);
    if (oleInitOk) {
        OleUninitialize();
    }
    if (comInitOk) {
        CoUninitialize();
    }
}

std::string HandleEditorOpen(const std::string& args, HostState& state) {
    bool isMusic = true;
    size_t index = 0;
    if (!ParseEditorArgs(args, isMusic, index)) {
        return "EDITOR_ERR reason=invalid_args";
    }
    std::string reason;
    if (!ValidateEditorIndex(state, isMusic, index, reason)) {
        return "EDITOR_ERR reason=" + reason;
    }
    UiCommand cmd{};
    cmd.type = UiCommandType::OpenEditor;
    cmd.isMusic = isMusic;
    cmd.index = index;
    EnqueueUiCommand(state, std::move(cmd));
    HostLog("editor_open queued chain=" + std::string(isMusic ? "music" : "mic") + " index=" + std::to_string(index));
    return std::string("EDITOR_OK queued=1 chain=") + (isMusic ? "music" : "mic") + " index=" + std::to_string(index);
}

std::string HandleEditorClose(const std::string& args, HostState& state) {
    bool isMusic = true;
    size_t index = 0;
    if (!ParseEditorArgs(args, isMusic, index)) {
        return "EDITOR_ERR reason=invalid_args";
    }
    std::string reason;
    if (!ValidateEditorIndex(state, isMusic, index, reason)) {
        return "EDITOR_ERR reason=" + reason;
    }
    UiCommand cmd{};
    cmd.type = UiCommandType::CloseEditor;
    cmd.isMusic = isMusic;
    cmd.index = index;
    EnqueueUiCommand(state, std::move(cmd));
    return std::string("EDITOR_OK queued=1 chain=") + (isMusic ? "music" : "mic") + " index=" + std::to_string(index) + " closed=1";
}

} // namespace

int wmain(int argc, wchar_t** argv) {
    std::wstring pipeName = L"micmix_vst_host";
    for (int i = 1; i + 1 < argc; ++i) {
        if (_wcsicmp(argv[i], L"--pipe") == 0) {
            pipeName = argv[i + 1];
        }
    }
    const std::wstring pipePath = std::wstring(L"\\\\.\\pipe\\") + pipeName;

    HostState state{};
    state.hostContext = new HostApplication();

    state.shmMap = CreateFileMappingW(
        INVALID_HANDLE_VALUE,
        nullptr,
        PAGE_READWRITE,
        0,
        static_cast<DWORD>(sizeof(micmix::vstipc::SharedMemory)),
        micmix::vstipc::kSharedMemoryName);
    if (state.shmMap) {
        void* view = MapViewOfFile(state.shmMap, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(micmix::vstipc::SharedMemory));
        state.shm = reinterpret_cast<micmix::vstipc::SharedMemory*>(view);
        if (state.shm &&
            (state.shm->magic != micmix::vstipc::kMagic || state.shm->version != micmix::vstipc::kVersion)) {
            micmix::vstipc::InitializeSharedMemory(*state.shm);
        }
        ResetAudioRings(state);
    }

    state.uiCmdThread = std::thread([&state]() { UiCommandMain(&state); });
    state.worker = std::thread([&state]() { WorkerMain(&state); });
    HostLog("host boot");

    bool running = true;
    while (running) {
        HANDLE pipe = CreateNamedPipeW(
            pipePath.c_str(),
            PIPE_ACCESS_DUPLEX,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
            1,
            static_cast<DWORD>(kMaxPipeMessageBytes),
            static_cast<DWORD>(kMaxPipeMessageBytes),
            300,
            nullptr);
        if (pipe == INVALID_HANDLE_VALUE) {
            Sleep(200);
            continue;
        }

        const BOOL connected = ConnectNamedPipe(pipe, nullptr) ? TRUE : (GetLastError() == ERROR_PIPE_CONNECTED);
        if (!connected) {
            CloseHandle(pipe);
            Sleep(50);
            continue;
        }

        std::string command;
        while (running && ReadMessage(pipe, command)) {
            if (_stricmp(command.c_str(), "PING") == 0 || _stricmp(command.c_str(), "STATUS") == 0) {
                WriteResponse(pipe, BuildPong(state));
                continue;
            }
            if (_stricmp(command.c_str(), "QUIT") == 0) {
                WriteResponse(pipe, "BYE");
                running = false;
                break;
            }
            if (command.size() > 5U && _strnicmp(command.c_str(), "SYNC ", 5) == 0) {
                WriteResponse(pipe, HandleSync(Trim(command.substr(5)), state));
                continue;
            }
            if (command.size() > 5U && _strnicmp(command.c_str(), "SCAN ", 5) == 0) {
                WriteResponse(pipe, HandleScan(Trim(command.substr(5))));
                continue;
            }
            if (command.size() > 12U && _strnicmp(command.c_str(), "EDITOR_OPEN ", 12) == 0) {
                WriteResponse(pipe, HandleEditorOpen(Trim(command.substr(12)), state));
                continue;
            }
            if (command.size() > 13U && _strnicmp(command.c_str(), "EDITOR_CLOSE ", 13) == 0) {
                WriteResponse(pipe, HandleEditorClose(Trim(command.substr(13)), state));
                continue;
            }
            WriteResponse(pipe, "ERR unknown_command");
        }

        FlushFileBuffers(pipe);
        DisconnectNamedPipe(pipe);
        CloseHandle(pipe);
    }

    state.workerRun.store(false, std::memory_order_release);
    if (state.worker.joinable()) {
        state.worker.join();
    }
    state.uiCmdRun.store(false, std::memory_order_release);
    UiCommand quitCmd{};
    quitCmd.type = UiCommandType::Quit;
    EnqueueUiCommand(state, std::move(quitCmd));
    if (state.uiCmdThread.joinable()) {
        state.uiCmdThread.join();
    }

    if (state.shm) {
        UnmapViewOfFile(state.shm);
        state.shm = nullptr;
    }
    if (state.shmMap) {
        CloseHandle(state.shmMap);
        state.shmMap = nullptr;
    }
    if (state.hostContext) {
        delete state.hostContext;
        state.hostContext = nullptr;
    }
    HostLog("host shutdown");
    return 0;
}
