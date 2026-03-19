#pragma once

class EffectsWindowController {
public:
    static EffectsWindowController& Instance();
    void Open();
    void Close();

private:
    EffectsWindowController() = default;
    ~EffectsWindowController() = default;
    EffectsWindowController(const EffectsWindowController&) = delete;
    EffectsWindowController& operator=(const EffectsWindowController&) = delete;
};

