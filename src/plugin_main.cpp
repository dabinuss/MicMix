#include <Windows.h>

#include <cassert>
#include <cstdlib>
#include <cstring>

#include "teamspeak/public_definitions.h"
#include "teamspeak/public_errors.h"
#include "plugin.h"
#include "plugin_definitions.h"
#include "micmix_core.h"

namespace {

constexpr int PLUGIN_API_VERSION = 26;
constexpr int MENU_ID_SETTINGS = 1;

char* g_pluginIdRaw = nullptr;

PluginMenuItem* CreateMenuItem(PluginMenuType type, int id, const char* text, const char* icon) {
    auto* item = static_cast<PluginMenuItem*>(std::malloc(sizeof(PluginMenuItem)));
    item->type = type;
    item->id = id;
    strcpy_s(item->text, PLUGIN_MENU_BUFSZ, text);
    strcpy_s(item->icon, PLUGIN_MENU_BUFSZ, icon);
    return item;
}

PluginHotkey* CreateHotkey(const char* keyword, const char* description) {
    auto* hotkey = static_cast<PluginHotkey*>(std::malloc(sizeof(PluginHotkey)));
    strcpy_s(hotkey->keyword, PLUGIN_HOTKEY_BUFSZ, keyword);
    strcpy_s(hotkey->description, PLUGIN_HOTKEY_BUFSZ, description);
    return hotkey;
}

} // namespace

extern "C" {

const char* ts3plugin_name() {
    return "MicMix";
}

const char* ts3plugin_version() {
    return "1.0.0";
}

int ts3plugin_apiVersion() {
    return PLUGIN_API_VERSION;
}

const char* ts3plugin_author() {
    return "MicMix";
}

const char* ts3plugin_description() {
    return "Mixes an additional audio source into captured microphone audio.";
}

void ts3plugin_setFunctionPointers(const struct TS3Functions funcs) {
    g_ts3Functions = funcs;
}

int ts3plugin_init() {
    char configPath[1024]{};
    g_ts3Functions.getConfigPath(configPath, sizeof(configPath));
    const bool ok = MicMixApp::Instance().Initialize(configPath);
    return ok ? 0 : 1;
}

void ts3plugin_shutdown() {
    MicMixApp::Instance().Shutdown();
    if (g_pluginIdRaw) {
        std::free(g_pluginIdRaw);
        g_pluginIdRaw = nullptr;
    }
}

void ts3plugin_registerPluginID(const char* id) {
    if (!id) {
        return;
    }
    const size_t len = std::strlen(id) + 1;
    g_pluginIdRaw = static_cast<char*>(std::malloc(len));
    strcpy_s(g_pluginIdRaw, len, id);
    g_pluginId = id;
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
    MicMixApp::Instance().OpenSettingsWindow();
}

void ts3plugin_freeMemory(void* data) {
    std::free(data);
}

void ts3plugin_initMenus(struct PluginMenuItem*** menuItems, char** menuIcon) {
    const size_t sz = 2;
    *menuItems = static_cast<PluginMenuItem**>(std::malloc(sizeof(PluginMenuItem*) * sz));
    (*menuItems)[0] = CreateMenuItem(PLUGIN_MENU_TYPE_GLOBAL, MENU_ID_SETTINGS, "MicMix Settings...", "");
    (*menuItems)[1] = nullptr;
    *menuIcon = nullptr;
}

void ts3plugin_initHotkeys(struct PluginHotkey*** hotkeys) {
    const size_t sz = 2;
    *hotkeys = static_cast<PluginHotkey**>(std::malloc(sizeof(PluginHotkey*) * sz));
    (*hotkeys)[0] = CreateHotkey("music_toggle_mute", "MicMix: Music mute (toggle)");
    (*hotkeys)[1] = nullptr;
}

void ts3plugin_onMenuItemEvent(uint64 serverConnectionHandlerID, enum PluginMenuType type, int menuItemID, uint64 selectedItemID) {
    (void)serverConnectionHandlerID;
    (void)selectedItemID;
    if (type == PLUGIN_MENU_TYPE_GLOBAL && menuItemID == MENU_ID_SETTINGS) {
        MicMixApp::Instance().OpenSettingsWindow();
    }
}

void ts3plugin_onHotkeyEvent(const char* keyword) {
    if (!keyword) return;
    if (std::strcmp(keyword, "music_toggle_mute") == 0) {
        MicMixApp::Instance().ToggleMute();
    }
}

void ts3plugin_currentServerConnectionChanged(uint64 serverConnectionHandlerID) {
    MicMixApp::Instance().SetActiveServer(serverConnectionHandlerID);
}

void ts3plugin_onTalkStatusChangeEvent(uint64 serverConnectionHandlerID, int status, int isReceivedWhisper, anyID clientID) {
    (void)isReceivedWhisper;
    MicMixApp::Instance().SetTalkStateForOwnClient(serverConnectionHandlerID, clientID, status);
}

void ts3plugin_onEditCapturedVoiceDataEvent(uint64 serverConnectionHandlerID, short* samples, int sampleCount, int channels, int* edited) {
    MicMixApp::Instance().EditCapturedVoice(serverConnectionHandlerID, samples, sampleCount, channels, edited);
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
    *data = nullptr;
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
