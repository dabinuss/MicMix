#pragma once

class SettingsWindowController {
public:
    static SettingsWindowController& Instance();
    void Open();
    void Close();

private:
    SettingsWindowController() = default;
    ~SettingsWindowController() = default;
    SettingsWindowController(const SettingsWindowController&) = delete;
    SettingsWindowController& operator=(const SettingsWindowController&) = delete;
};
