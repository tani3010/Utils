# -*- coding: utf-8 -*-
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd

class MyPlot:
    def __init__(self):
        self.initialize_plot_setting_before()

    def __del__(self):
        pass

    @staticmethod
    def initialize_plot_setting_before(
        xtick_direction='in', ytick_direction='in',
        xtick_major_size=3.5, ytick_major_size=3.5,
        xtick_minor_size=2.0, ytick_minor_size=2.0):
        # plt.style.use('seaborn-pastel')
        # plt.style.use('seaborn-whitegrid')
        plt.rcParams['xtick.direction'] = xtick_direction
        plt.rcParams['ytick.direction'] = ytick_direction
        plt.rcParams['xtick.major.size'] = xtick_major_size
        plt.rcParams['ytick.major.size'] = ytick_major_size
        plt.rcParams['xtick.minor.size'] = xtick_minor_size
        plt.rcParams['ytick.minor.size'] = ytick_minor_size

    @staticmethod
    def initialize_plot_setting_after():
        plt.tight_layout()
        plt.grid(linestyle='--', alpha=0.4)
        plt.legend()

    @staticmethod
    def set_main_title(title=None):
        if title is not None:
            plt.title(title)

    @staticmethod
    def set_xlabel_ylabel(xlabel, ylabel):
        if xlabel is not None:
            plt.xlabel(xlabel)

        if ylabel is not None:
            plt.ylabel(ylabel)

    @staticmethod
    def set_xlim_ylim(xlim, ylim):
        if xlim is not None:
            plt.xlim(xlim)

        if ylim is not None:
            plt.ylim(ylim)

    @staticmethod
    def save_figure(file_name=None):
        if file_name is not None:
            plt.savefig(file_name)

    @staticmethod
    def rotate_xticks_label(angle=90):
        plt.xticks(rotation=angle)

    def plot_chart(self, title, xlabel, ylabel, xdata, ydata, labels,
                   xticks_label_angle=0, xlim=None, ylim=None,
                   save_file_name=None):
        self.initialize_plot_setting_before()
        plt.figure()
        self.set_main_title(title)
        self.set_xlim_ylim(xlim, ylim)

        for x, y, l in zip(xdata, ydata, labels):
            plt.plot(x, y, label=l)

        self.set_xlabel_ylabel(xlabel, ylabel)
        self.rotate_xticks_label(xticks_label_angle)
        self.initialize_plot_setting_after()
        self.save_figure(save_file_name)
        plt.show()

    def plot_chart_with_errorbar(
        self, title, xlabel, ylabel, xdata, ydata, yerr, labels,
        xticks_label_angle=0, xlim=None, ylim=None, save_file_name=None):
        self.initialize_plot_setting_before()
        fig = plt.figure()
        ax = fig.add_subplot(1, 1, 1)
        self.set_main_title(title)
        self.set_xlim_ylim(xlim, ylim)

        for x, y, ye, l in zip(xdata, ydata, yerr, labels):
            ax.errorbar(x, y, yerr=ye, label=l, capsize=3)

        self.set_xlabel_ylabel(xlabel, ylabel)
        self.rotate_xticks_label(xticks_label_angle)
        self.initialize_plot_setting_after()
        self.save_figure(save_file_name)
        plt.show()

    def plot_barchart(
        self, title, xlabel, ylabel, xdata, ydata, labels,
        xticks_label_angle=90, xlim=None, ylim=None, save_file_name=None):
        self.initialize_plot_setting_before(xtick_major_size=0)
        fig = plt.figure()
        ax = fig.add_subplot(1, 1, 1)
        self.set_main_title(title)
        self.set_xlim_ylim(xlim, ylim)

        width = 1/len(labels) - 0.03
        count = 0
        xposition = np.arange(len(xdata))
        for y, l in zip(ydata, labels):
            ax.bar(xposition+width*count, y, width=width, label=l)
            count += 1

        ax.set_xticks(xposition+width*(count-1)*0.5)
        ax.set_xticklabels(xdata)

        self.set_xlabel_ylabel(xlabel, ylabel)
        self.rotate_xticks_label(xticks_label_angle)
        self.initialize_plot_setting_after()
        self.save_figure(save_file_name)
        plt.show()

    def plot_box_and_violin_plot(
        self, title, xlabel, ylabel, xdata, ydata, labels, is_violin=True,
        xticks_label_angle=90, xlim=None, ylim=None, save_file_name=None):
        self.initialize_plot_setting_before(xtick_major_size=0)
        self.set_main_title(title)
        self.set_xlim_ylim(xlim, ylim)

        df = pd.DataFrame(columns=['xdata', 'ydata', 'label'])
        for i in range(len(ydata)):
            for j in range(len(labels)):
                tmp_df = pd.DataFrame(
                    {
                        'xdata': xdata[i],
                        'ydata': ydata[i][j],
                        'label': labels[j]
                    }
                )
                df = df.append(tmp_df, ignore_index=True)

        df = df.astype({'xdata': 'category', 'ydata': 'float64', 'label': 'category'})
        if is_violin:
            sns.violinplot(x="xdata", y="ydata", hue="label", data=df)
        else:
            sns.boxplot(x="xdata", y="ydata", hue="label", data=df)

        self.set_xlabel_ylabel(xlabel, ylabel)
        self.rotate_xticks_label(xticks_label_angle)
        self.initialize_plot_setting_after()
        self.save_figure(save_file_name)
        plt.show()


if __name__  ==  '__main__':
    labels = ['XXX', 'YYY', 'ZZZ', 'AAA', 'BBB']
    xdata = ['AAA', 'BBB', 'CCC', 'DDD', 'EEE']
    dat = [10, 20, 30, 40, 50]
    dat2 = [11, 21, 31, 41, 50]
    dat3 = [12, 13, 33, 45, 55]
    dat4 = [dat, dat2, dat3, dat, dat3]
    yerr = np.linspace(1, 5, 5)

    my_plot = MyPlot()
    my_plot.plot_chart(
        'Test Chart', 'x_data', 'y_data',
        [xdata, xdata, xdata],
        [dat, dat2, dat3],
        labels,
        xticks_label_angle=90)

    my_plot.plot_chart_with_errorbar(
        'Test Chart with error bar', 'x_data', 'y_data',
        [xdata, xdata, xdata],
        [dat, dat2, dat3],
        [yerr, yerr, yerr],
        labels,
        xticks_label_angle=0)

    my_plot.plot_barchart(
        'Test grouped barplot', 'x_data', 'y_data',
        xdata, [dat, dat2, dat3, dat2, dat2], labels,
        xticks_label_angle=0)

    my_plot.plot_box_and_violin_plot(
        'Test grouped barplot', 'x_data', 'y_data',
        xdata, [dat4, dat4, dat4, dat4, dat4], labels, True,
        xticks_label_angle=0)

    my_plot.plot_box_and_violin_plot(
        'Test grouped barplot', 'x_data', 'y_data',
        xdata, [dat4, dat4, dat4, dat4, dat4], labels, False,
        xticks_label_angle=0)
