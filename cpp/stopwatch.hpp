#ifndef __STOPWATCH_HPP__
#define __STOPWATCH_HPP__

#include <chrono>
#include <iostream>
#include <iomanip>

template <typename RESOLUTION = std::chrono::seconds>
class stopWatch {
public:
    stopWatch(const unsigned int digits_ = 3) : 
        start(std::chrono::system_clock::now()),
        elapsed(0.0),
        digits(digits_),
        unit(getResolutionString()) {}
    ~stopWatch() {}
    inline void show() {
        end = std::chrono::system_clock::now();
        elapsed = std::chrono::duration_cast<RESOLUTION>(end - start).count();
        std::cout << " elapsed: " << std::fixed << std::setprecision(digits_)
          << elapsed << "[" << unit << "]" << std::endl;
    }
    inline void reset() { start = std::chrono::system_clock::now(); }
private:
    const int digits;
    std::chrono::system_clock::time_point start, end;
    double elapsed;
    std::string unit;
    std::string getResolutionString() const {
        if constexpr (std::is_same_v<RESOLUTION, std::chrono::nanoseconds>) return "ns";
        if constexpr (std::is_same_v<RESOLUTION, std::chrono::microseconds>) return "us";
        if constexpr (std::is_same_v<RESOLUTION, std::chrono::milliseconds>) return "ms";
        if constexpr (std::is_same_v<RESOLUTION, std::chrono::seconds>) return "s";
        if constexpr (std::is_same_v<RESOLUTION, std::chrono::minutes>) return "m";
        if constexpr (std::is_same_v<RESOLUTION, std::chrono::hours>) return "h";
        return "unknown";
    }
};

#endif  // __STOPWATCH_HPP__
