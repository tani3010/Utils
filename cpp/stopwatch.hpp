#ifndef __STOPWATCH_HPP__
#define __STOPWATCH_HPP__

#include <chrono>
#include <iomanip>
#include <iostream>

template <typename RESOLUTION = std::chrono::duration<float> >
class stopWatch {
public:
  stopWatch() : start(std::chrono::system_clock::now()), elapsed(0.0), digits(3) {}
  ~stopWatch() {}
  inline void show(const std::string& unit = "s") {
    end = std::chrono::system_clock::now();
    elapsed = std::chrono::duration_cast<RESOLUTION>(end - start).count();
    std::cout << " elapsed: " << std::fixed << std::setprecision(digits)
      << elapsed << "[" << unit << "]" << std::endl;
  }
  inline void reset() { start = std::chrono::system_clock::now(); }
private:
  const int digits;
  std::chrono::system_clock::time_point start, end;
  double elapsed;
};

#endif // __STOPWATCH_HPP__