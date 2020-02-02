import QuantLib as ql
import numpy as np
import matplotlib.pyplot as plt

def interpolate(x, y, target, type):
    funct = type(x, y)
    return map(lambda target_: funct(target_, True), target)

if __name__ == '__main__':
    dict = {"BackwardFlatInterpolation"      : ql.BackwardFlatInterpolation,
            "ForwardFlatInterpolation"       : ql.ForwardFlatInterpolation,
            "LinearInterpolation"            : ql.LinearInterpolation,
            "LogLinearInterpolation"         : ql.LogLinearInterpolation,
            "CubicNaturalSpline"             : ql.CubicNaturalSpline,
            "LogCubicNaturalSpline"          : ql.LogCubicNaturalSpline,
            "MonotonicCubicNaturalSpline"    : ql.MonotonicCubicNaturalSpline,
            "MonotonicLogCubicNaturalSpline" : ql.MonotonicLogCubicNaturalSpline,
            "KrugerCubic"                    : ql.KrugerCubic,
            "KrugerLogCubic"                 : ql.KrugerLogCubic,
            "Parabolic"                      : ql.Parabolic,
            "LogParabolic"                   : ql.LogParabolic,
            "MonotonicParabolic"             : ql.MonotonicParabolic,
            "MonotonicLogParabolic"          : ql.MonotonicLogParabolic,
            "FritschButlandCubic"            : ql.FritschButlandCubic,
            "FritschButlandLogCubic"         : ql.FritschButlandLogCubic}
    x = np.linspace(1, 10, 10).tolist()
    y = (np.sin(x) + 10).tolist()
    target = np.linspace(1, 11, 200).tolist()
    plt.plot(x, y, "ro", markersize=8)
    for key in dict.keys():
        plt.plot(target, interpolate(x, y, target, dict[key]))
    label = dict.keys(); label.insert(0, "raw data")
    plt.legend(label, "lower left")