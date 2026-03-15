#include <Windows.h>

#include <cassert>
#include <atomic>
#include <cstdlib>
#include <cstring>
#include <exception>
#include <string>

#include "teamspeak/public_definitions.h"
#include "teamspeak/public_errors.h"
#include "plugin.h"
#include "plugin_definitions.h"
#include "micmix_core.h"
#include "micmix_version.h"

namespace {

constexpr int PLUGIN_API_VERSION = 26;
constexpr int MENU_ID_SETTINGS = 1;
constexpr char MENU_ICON_FILE[] = "t.png";
constexpr char SETTINGS_ICON_FILE[] = "1.png";

char* g_pluginIdRaw = nullptr;
std::atomic<bool> g_pluginShuttingDown{false};

bool IsPluginOperational() {
    return !g_pluginShuttingDown.load(std::memory_order_acquire) &&
           MicMixApp::Instance().IsInitialized();
}

template <typename Fn>
void RunGuarded(const char* where, Fn&& fn) {
    try {
        fn();
    } catch (const std::exception& ex) {
        if (!g_pluginShuttingDown.load(std::memory_order_acquire)) {
            LogError(std::string(where) + " exception: " + ex.what());
        }
    } catch (...) {
        if (!g_pluginShuttingDown.load(std::memory_order_acquire)) {
            LogError(std::string(where) + " exception: unknown");
        }
    }
}

PluginMenuItem* CreateMenuItem(PluginMenuType type, int id, const char* text, const char* icon) {
    auto* item = static_cast<PluginMenuItem*>(std::calloc(1, sizeof(PluginMenuItem)));
    if (!item) {
        return nullptr;
    }
    item->type = type;
    item->id = id;
    strcpy_s(item->text, PLUGIN_MENU_BUFSZ, text ? text : "");
    strcpy_s(item->icon, PLUGIN_MENU_BUFSZ, icon ? icon : "");
    return item;
}

PluginHotkey* CreateHotkey(const char* keyword, const char* description) {
    auto* hotkey = static_cast<PluginHotkey*>(std::calloc(1, sizeof(PluginHotkey)));
    if (!hotkey) {
        return nullptr;
    }
    strcpy_s(hotkey->keyword, PLUGIN_HOTKEY_BUFSZ, keyword ? keyword : "");
    strcpy_s(hotkey->description, PLUGIN_HOTKEY_BUFSZ, description ? description : "");
    return hotkey;
}

} // namespace

extern "C" {

const char* ts3plugin_name() {
    return "MicMix";
}

const char* ts3plugin_version() {
    return MICMIX_VERSION;
}

int ts3plugin_apiVersion() {
    return PLUGIN_API_VERSION;
}

const char* ts3plugin_author() {
    return "Dabinuss";
}

const char* ts3plugin_description() {
    return "Mixes an additional audio source into captured microphone audio.";
}

void ts3plugin_setFunctionPointers(const struct TS3Functions funcs) {
    g_pluginShuttingDown.store(false, std::memory_order_release);
    g_ts3Functions = funcs;
    SetTsLoggingEnabled(true);
}

int ts3plugin_init() {
    try {
        g_pluginShuttingDown.store(false, std::memory_order_release);
        if (!g_ts3Functions.getConfigPath) {
            return 1;
        }
        char configPath[1024]{};
        g_ts3Functions.getConfigPath(configPath, sizeof(configPath));
        configPath[sizeof(configPath) - 1] = '\0';
        if (configPath[0] == '\0') {
            return 1;
        }
        const bool ok = MicMixApp::Instance().Initialize(configPath);
        return ok ? 0 : 1;
    } catch (const std::exception& ex) {
        LogError(std::string("ts3plugin_init exception: ") + ex.what());
        return 1;
    } catch (...) {
        LogError("ts3plugin_init exception: unknown");
        return 1;
    }
}

void ts3plugin_shutdown() {
    g_pluginShuttingDown.store(true, std::memory_order_release);
    SetTsLoggingEnabled(false);
    RunGuarded("ts3plugin_shutdown", [] {
        MicMixApp::Instance().Shutdown();
    });
    if (g_pluginIdRaw) {
        std::free(g_pluginIdRaw);
        g_pluginIdRaw = nullptr;
    }
    g_pluginId.clear();
}

void ts3plugin_registerPluginID(const char* id) {
    if (!id || id[0] == '\0') {
        return;
    }
    constexpr size_t kMaxPluginIdLen = 1024;
    const size_t idLen = strnlen_s(id, kMaxPluginIdLen + 1);
    if (idLen == 0 || idLen > kMaxPluginIdLen) {
        return;
    }
    if (g_pluginIdRaw) {
        std::free(g_pluginIdRaw);
        g_pluginIdRaw = nullptr;
    }
    const size_t len = idLen + 1;
    g_pluginIdRaw = static_cast<char*>(std::malloc(len));
    if (!g_pluginIdRaw) {
        return;
    }
    strcpy_s(g_pluginIdRaw, len, id);
    // SDK note: the incoming id pointer is only valid during this callback.
    g_pluginId = g_pluginIdRaw;
}

int ts3plugin_requestAutoload() {
    return 1;
}

int ts3plugin_offersConfigure() {
    return PLUGIN_OFFERS_CONFIGURE_NEW_THREAD;
}

void ts3plugin_configure(void* handle, void* qParentWidget) {
    (void)handle;
    (void)qParentWidget;
    RunGuarded("ts3plugin_configure", [] {
        MicMixApp::Instance().OpenSettingsWindow();
    });
}

void ts3plugin_freeMemory(void* data) {
    std::free(data);
}

void ts3plugin_initMenus(struct PluginMenuItem*** menuItems, char** menuIcon) {
    if (!menuItems || !menuIcon) {
        return;
    }
    *menuItems = nullptr;
    *menuIcon = nullptr;
    const size_t sz = 2;
    auto** items = static_cast<PluginMenuItem**>(std::calloc(sz, sizeof(PluginMenuItem*)));
    if (!items) {
        return;
    }
    items[0] = CreateMenuItem(PLUGIN_MENU_TYPE_GLOBAL, MENU_ID_SETTINGS, "MicMix Settings...", SETTINGS_ICON_FILE);
    if (!items[0]) {
        std::free(items);
        return;
    }
    auto* icon = static_cast<char*>(std::malloc(PLUGIN_MENU_BUFSZ));
    if (!icon) {
        std::free(items[0]);
        std::free(items);
        return;
    }
    strcpy_s(icon, PLUGIN_MENU_BUFSZ, MENU_ICON_FILE);
    *menuItems = items;
    *menuIcon = icon;
}

void ts3plugin_initHotkeys(struct PluginHotkey*** hotkeys) {
    if (!hotkeys) {
        return;
    }
    *hotkeys = nullptr;
    const size_t sz = 6;
    auto** keys = static_cast<PluginHotkey**>(std::calloc(sz, sizeof(PluginHotkey*)));
    if (!keys) {
        return;
    }
    auto releaseKeys = [keys, sz]() {
        for (size_t i = 0; i < sz; ++i) {
            if (keys[i]) {
                std::free(keys[i]);
                keys[i] = nullptr;
            }
        }
        std::free(keys);
    };

    keys[0] = CreateHotkey("music_toggle_mute", "MicMix: Music mute (toggle)");
    keys[1] = CreateHotkey("mic_input_toggle_mute", "MicMix: Microphone mute (toggle)");
    keys[2] = CreateHotkey("music_toggle_push_to_play", "MicMix: Push-to-play mode (toggle)");
    keys[3] = CreateHotkey("music_push_to_play_on", "MicMix: Push-to-play mode on");
    keys[4] = CreateHotkey("music_push_to_play_off", "MicMix: Push-to-play mode off");
    if (!keys[0] || !keys[1] || !keys[2] || !keys[3] || !keys[4]) {
        releaseKeys();
        return;
    }
    *hotkeys = keys;
}

void ts3plugin_onMenuItemEvent(uint64 serverConnectionHandlerID, enum PluginMenuType type, int menuItemID, uint64 selectedItemID) {
    if (!IsPluginOperational()) {
        return;
    }
    RunGuarded("ts3plugin_onMenuItemEvent", [&] {
        (void)serverConnectionHandlerID;
        (void)selectedItemID;
        if (type == PLUGIN_MENU_TYPE_GLOBAL && menuItemID == MENU_ID_SETTINGS) {
            MicMixApp::Instance().OpenSettingsWindow();
        }
    });
}

void ts3plugin_onHotkeyEvent(const char* keyword) {
    if (!IsPluginOperational()) {
        return;
    }
    RunGuarded("ts3plugin_onHotkeyEvent", [&] {
        if (!keyword) return;
        if (std::strcmp(keyword, "music_toggle_mute") == 0) {
            MicMixApp::Instance().ToggleMute();
            return;
        }
        if (std::strcmp(keyword, "mic_input_toggle_mute") == 0) {
            MicMixApp::Instance().ToggleMicInputMute();
            return;
        }
        if (std::strcmp(keyword, "music_toggle_push_to_play") == 0) {
            MicMixApp::Instance().TogglePushToPlay();
            return;
        }
        if (std::strcmp(keyword, "music_push_to_play_on") == 0) {
            MicMixApp::Instance().SetPushToPlayActive(true);
            return;
        }
        if (std::strcmp(keyword, "music_push_to_play_off") == 0) {
            MicMixApp::Instance().SetPushToPlayActive(false);
        }
    });
}

void ts3plugin_currentServerConnectionChanged(uint64 serverConnectionHandlerID) {
    if (!IsPluginOperational()) {
        return;
    }
    RunGuarded("ts3plugin_currentServerConnectionChanged", [&] {
        MicMixApp::Instance().SetActiveServer(serverConnectionHandlerID);
    });
}

void ts3plugin_onTalkStatusChangeEvent(uint64 serverConnectionHandlerID, int status, int isReceivedWhisper, anyID clientID) {
    if (!IsPluginOperational()) {
        return;
    }
    RunGuarded("ts3plugin_onTalkStatusChangeEvent", [&] {
        (void)isReceivedWhisper;
        MicMixApp::Instance().SetTalkStateForOwnClient(serverConnectionHandlerID, clientID, status);
    });
}

void ts3plugin_onConnectStatusChangeEvent(uint64 serverConnectionHandlerID, int newStatus, unsigned int errorNumber) {
    if (!IsPluginOperational()) {
        return;
    }
    RunGuarded("ts3plugin_onConnectStatusChangeEvent", [&] {
        MicMixApp::Instance().OnConnectStatusChange(serverConnectionHandlerID, newStatus, errorNumber);
    });
}

void ts3plugin_onEditCapturedVoiceDataEvent(uint64 serverConnectionHandlerID, short* samples, int sampleCount, int channels, int* edited) {
    if (!IsPluginOperational()) {
        return;
    }
    RunGuarded("ts3plugin_onEditCapturedVoiceDataEvent", [&] {
        MicMixApp::Instance().EditCapturedVoice(serverConnectionHandlerID, samples, sampleCount, channels, edited);
    });
}

const char* ts3plugin_commandKeyword() {
    return "";
}

int ts3plugin_processCommand(uint64 serverConnectionHandlerID, const char* command) {
    (void)serverConnectionHandlerID;
    (void)command;
    return 1;
}

const char* ts3plugin_infoTitle() {
    return "MicMix";
}

void ts3plugin_infoData(uint64 serverConnectionHandlerID, uint64 id, PluginItemType type, char** data) {
    (void)serverConnectionHandlerID;
    (void)id;
    (void)type;
    if (data) {
        *data = nullptr;
    }
}

const char* ts3plugin_keyPrefix() {
    return "micmix";
}

const char* ts3plugin_keyDeviceName(const char* keyIdentifier) {
    (void)keyIdentifier;
    return nullptr;
}

const char* ts3plugin_displayKeyText(const char* keyIdentifier) {
    (void)keyIdentifier;
    return nullptr;
}

} // extern "C"
