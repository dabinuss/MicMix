#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <algorithm>
#include <cstdint>
#include <cstring>

namespace micmix::vstipc {

constexpr uint32_t kMagic = 0x56535043U; // 'VSPC'
constexpr uint32_t kVersion = 2U;
constexpr uint32_t kRingCapacity = 128U;
constexpr uint32_t kMaxFramesPerPacket = 960U;
constexpr wchar_t kSharedMemoryName[] = L"Local\\MicMixVstAudioV2";

struct AudioPacket {
    uint32_t seq = 0;
    uint32_t frames = 0;
    uint32_t channels = 1;
    uint32_t reserved = 0;
    float samples[kMaxFramesPerPacket]{};
};

struct alignas(8) AudioRing {
    volatile LONGLONG readIndex = 0;
    volatile LONGLONG writeIndex = 0;
    AudioPacket packets[kRingCapacity]{};
};

struct SharedMemory {
    uint32_t magic = kMagic;
    uint32_t version = kVersion;
    volatile LONG hostHeartbeat = 0;
    volatile LONG pluginHeartbeat = 0;
    AudioRing musicIn{};
    AudioRing musicOut{};
    AudioRing micIn{};
    AudioRing micOut{};
};

inline LONG AtomicLoad(const volatile LONG* value) {
    return InterlockedCompareExchange(const_cast<volatile LONG*>(value), 0, 0);
}

inline LONGLONG AtomicLoad64(const volatile LONGLONG* value) {
    return InterlockedCompareExchange64(const_cast<volatile LONGLONG*>(value), 0, 0);
}

inline bool RingPush(AudioRing& ring, const AudioPacket& packet) {
    LONGLONG write = AtomicLoad64(&ring.writeIndex);
    const LONGLONG read = AtomicLoad64(&ring.readIndex);
    if ((write - read) >= static_cast<LONGLONG>(kRingCapacity)) {
        return false;
    }
    AudioPacket bounded = packet;
    bounded.frames = std::min<uint32_t>(bounded.frames, kMaxFramesPerPacket);
    bounded.channels = std::max<uint32_t>(1U, bounded.channels);
    ring.packets[static_cast<size_t>(write) % kRingCapacity] = bounded;
    MemoryBarrier();
    InterlockedExchange64(&ring.writeIndex, write + 1);
    return true;
}

inline bool RingPop(AudioRing& ring, AudioPacket& out) {
    LONGLONG read = AtomicLoad64(&ring.readIndex);
    const LONGLONG write = AtomicLoad64(&ring.writeIndex);
    if (read >= write) {
        return false;
    }
    out = ring.packets[static_cast<size_t>(read) % kRingCapacity];
    MemoryBarrier();
    InterlockedExchange64(&ring.readIndex, read + 1);
    return true;
}

inline void RingReset(AudioRing& ring) {
    const LONGLONG write = AtomicLoad64(&ring.writeIndex);
    InterlockedExchange64(&ring.readIndex, write);
}

inline void InitializeSharedMemory(SharedMemory& shm) {
    std::memset(&shm, 0, sizeof(SharedMemory));
    shm.magic = kMagic;
    shm.version = kVersion;
}

} // namespace micmix::vstipc
