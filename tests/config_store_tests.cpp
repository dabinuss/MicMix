#include <chrono>
#include <cmath>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <string>
#include <vector>

#include "micmix_core.h"

namespace {

namespace fs = std::filesystem;
std::vector<fs::path> g_tempDirs;

bool NearlyEqual(float a, float b, float eps = 0.0001f) {
    return std::fabs(a - b) <= eps;
}

fs::path MakeUniqueBaseDir(const char* testName) {
    const auto tick = std::chrono::high_resolution_clock::now().time_since_epoch().count();
    const fs::path base = fs::temp_directory_path() /
        ("micmix_test_" + std::string(testName) + "_" + std::to_string(tick));
    std::error_code ec;
    fs::create_directories(base, ec);
    g_tempDirs.push_back(base);
    return base;
}

void CleanupTempDirs() {
    for (const auto& dir : g_tempDirs) {
        std::error_code ec;
        fs::remove_all(dir, ec);
    }
    g_tempDirs.clear();
}

bool WriteFile(const std::string& path, const std::string& payload) {
    std::ofstream out(path, std::ios::binary | std::ios::trunc);
    if (!out.is_open()) {
        return false;
    }
    out.write(payload.data(), static_cast<std::streamsize>(payload.size()));
    out.flush();
    return out.good();
}

bool Expect(bool cond, const std::string& msg) {
    if (!cond) {
        std::cerr << "FAIL: " << msg << std::endl;
        return false;
    }
    return true;
}

bool TestJsonWithEqualsInString() {
    const fs::path base = MakeUniqueBaseDir("json_equals");
    ConfigStore store(base.string());
    const std::string payload =
        "{\n"
        "  \"source.mode\": \"app_session\",\n"
        "  \"source.app.process_name\": \"foo=bar.exe\",\n"
        "  \"mix.music_gain_db\": -12.5,\n"
        "  \"source.autostart_enabled\": true\n"
        "}\n";
    if (!WriteFile(store.ConfigPath(), payload)) {
        std::cerr << "FAIL: could not write config file for json_equals" << std::endl;
        return false;
    }

    MicMixSettings settings{};
    settings.sourceMode = SourceMode::Loopback;
    settings.appProcessName = "default.exe";
    settings.musicGainDb = -20.0f;
    settings.autostartEnabled = false;

    std::string warning;
    const bool ok = store.Load(settings, warning);
    bool pass = true;
    pass &= Expect(ok, "Load(json_equals) should return true");
    pass &= Expect(settings.sourceMode == SourceMode::AppSession, "JSON source.mode should be parsed as app_session");
    pass &= Expect(settings.appProcessName == "foo=bar.exe", "JSON string with '=' should not be treated as INI pair");
    pass &= Expect(NearlyEqual(settings.musicGainDb, -12.5f), "JSON float value should be parsed");
    pass &= Expect(settings.autostartEnabled, "JSON bool value should be parsed");
    pass &= Expect(warning.empty(), "Valid JSON config should not produce parse warnings");
    return pass;
}

bool TestIniStillParses() {
    const fs::path base = MakeUniqueBaseDir("ini_ok");
    ConfigStore store(base.string());
    const std::string payload =
        "source.mode=app_session\n"
        "source.app.process_name=spotify.exe\n"
        "mix.music_gain_db=-18.0\n"
        "source.autostart_enabled=true\n";
    if (!WriteFile(store.ConfigPath(), payload)) {
        std::cerr << "FAIL: could not write config file for ini_ok" << std::endl;
        return false;
    }

    MicMixSettings settings{};
    std::string warning;
    const bool ok = store.Load(settings, warning);
    bool pass = true;
    pass &= Expect(ok, "Load(ini_ok) should return true");
    pass &= Expect(settings.sourceMode == SourceMode::AppSession, "INI source.mode should be parsed");
    pass &= Expect(settings.appProcessName == "spotify.exe", "INI process name should be parsed");
    pass &= Expect(NearlyEqual(settings.musicGainDb, -18.0f), "INI music gain should be parsed");
    pass &= Expect(settings.autostartEnabled, "INI autostart should be parsed");
    pass &= Expect(warning.empty(), "Valid INI config should not produce parse warnings");
    return pass;
}

bool TestInvalidJsonReportsParseWarning() {
    const fs::path base = MakeUniqueBaseDir("invalid_json");
    ConfigStore store(base.string());
    const std::string payload =
        "{\n"
        "  \"source.mode\": \"app_session\",\n"
        "  \"source.app.process_name\": \"foo=bar.exe\"\n";
    if (!WriteFile(store.ConfigPath(), payload)) {
        std::cerr << "FAIL: could not write config file for invalid_json" << std::endl;
        return false;
    }

    MicMixSettings settings{};
    settings.sourceMode = SourceMode::Loopback;
    settings.appProcessName = "fallback.exe";

    std::string warning;
    const bool ok = store.Load(settings, warning);
    bool pass = true;
    pass &= Expect(ok, "Load(invalid_json) should return true with warning");
    pass &= Expect(settings.sourceMode == SourceMode::Loopback, "Invalid JSON must not be silently interpreted as INI");
    pass &= Expect(settings.appProcessName == "fallback.exe", "Invalid JSON should keep fallback process name");
    pass &= Expect(warning.find("Config parse issue detected") != std::string::npos,
                   "Invalid JSON should emit parse warning");
    return pass;
}

bool TestOversizedConfigIgnored() {
    const fs::path base = MakeUniqueBaseDir("oversized_config");
    ConfigStore store(base.string());
    std::string payload(1024 * 1024 + 64, 'a');
    payload += "\nsource.mode=app_session\n";
    if (!WriteFile(store.ConfigPath(), payload)) {
        std::cerr << "FAIL: could not write config file for oversized_config" << std::endl;
        return false;
    }

    MicMixSettings settings{};
    settings.sourceMode = SourceMode::Loopback;

    std::string warning;
    const bool ok = store.Load(settings, warning);
    bool pass = true;
    pass &= Expect(ok, "Load(oversized_config) should return true with warning");
    pass &= Expect(settings.sourceMode == SourceMode::Loopback, "Oversized config should be ignored");
    pass &= Expect(warning.find("Config too large; file ignored.") != std::string::npos,
                   "Oversized config should emit size warning");
    return pass;
}

bool TestVstEffectsIniParse() {
    const fs::path base = MakeUniqueBaseDir("vst_ini");
    ConfigStore store(base.string());
    const std::string payload =
        "vst.effects_enabled=true\n"
        "vst.host.autostart=true\n"
        "vst.music.count=2\n"
        "vst.music.0.path=C:\\\\VST\\\\Comp.vst3\n"
        "vst.music.0.name=Compressor\n"
        "vst.music.0.enabled=true\n"
        "vst.music.0.bypass=false\n"
        "vst.music.1.path=C:\\\\VST\\\\EQ.vst3\n"
        "vst.music.1.name=EQ\n"
        "vst.music.1.enabled=true\n"
        "vst.music.1.bypass=true\n"
        "vst.mic.count=1\n"
        "vst.mic.0.path=C:\\\\VST\\\\Gate.vst3\n"
        "vst.mic.0.name=Gate\n"
        "vst.mic.0.enabled=false\n"
        "vst.mic.0.bypass=true\n";
    if (!WriteFile(store.ConfigPath(), payload)) {
        std::cerr << "FAIL: could not write config file for vst_ini" << std::endl;
        return false;
    }

    MicMixSettings settings{};
    std::string warning;
    const bool ok = store.Load(settings, warning);
    bool pass = true;
    pass &= Expect(ok, "Load(vst_ini) should return true");
    pass &= Expect(settings.vstEffectsEnabled, "vst.effects_enabled should parse");
    pass &= Expect(settings.vstHostAutostart, "vst.host.autostart should parse");
    pass &= Expect(settings.musicEffects.size() == 2, "vst.music.count should parse");
    pass &= Expect(settings.micEffects.size() == 1, "vst.mic.count should parse");
    pass &= Expect(settings.musicEffects[0].name == "Compressor", "music effect name should parse");
    pass &= Expect(settings.musicEffects[1].bypass, "music bypass should parse");
    pass &= Expect(!settings.micEffects[0].enabled, "mic effect enabled should parse");
    pass &= Expect(warning.empty(), "Valid VST INI config should not produce parse warnings");
    return pass;
}

bool TestVstEffectsSaveRoundtrip() {
    const fs::path base = MakeUniqueBaseDir("vst_roundtrip");
    ConfigStore store(base.string());

    MicMixSettings writeSettings{};
    writeSettings.vstEffectsEnabled = true;
    writeSettings.vstHostAutostart = false;
    writeSettings.musicEffects = {
        { "C:\\\\VST\\\\Comp.vst3", "Compressor", true, false },
        { "C:\\\\VST\\\\EQ.vst3", "EQ", true, true }
    };
    writeSettings.micEffects = {
        { "C:\\\\VST\\\\Gate.vst3", "Gate", false, true }
    };

    std::string saveErr;
    const bool saveOk = store.Save(writeSettings, saveErr);
    bool pass = true;
    pass &= Expect(saveOk, "Save(vst_roundtrip) should return true");
    pass &= Expect(saveErr.empty(), "Save(vst_roundtrip) should not return errors");

    MicMixSettings readSettings{};
    std::string warning;
    const bool loadOk = store.Load(readSettings, warning);
    pass &= Expect(loadOk, "Load(vst_roundtrip) should return true");
    pass &= Expect(readSettings.vstEffectsEnabled, "roundtrip effects enabled should match");
    pass &= Expect(!readSettings.vstHostAutostart, "roundtrip host autostart should match");
    pass &= Expect(readSettings.musicEffects.size() == 2, "roundtrip music effect count should match");
    pass &= Expect(readSettings.micEffects.size() == 1, "roundtrip mic effect count should match");
    pass &= Expect(readSettings.musicEffects[1].bypass, "roundtrip music bypass should match");
    pass &= Expect(!readSettings.micEffects[0].enabled, "roundtrip mic enabled should match");
    pass &= Expect(warning.empty(), "Roundtrip load should not produce parse warnings");
    return pass;
}

} // namespace

int main() {
    int failed = 0;
    if (!TestJsonWithEqualsInString()) {
        ++failed;
    }
    if (!TestIniStillParses()) {
        ++failed;
    }
    if (!TestInvalidJsonReportsParseWarning()) {
        ++failed;
    }
    if (!TestOversizedConfigIgnored()) {
        ++failed;
    }
    if (!TestVstEffectsIniParse()) {
        ++failed;
    }
    if (!TestVstEffectsSaveRoundtrip()) {
        ++failed;
    }
    CleanupTempDirs();

    if (failed != 0) {
        std::cerr << "config_store_tests failed: " << failed << std::endl;
        return 1;
    }
    std::cout << "config_store_tests passed" << std::endl;
    return 0;
}
