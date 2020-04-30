# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np

class DataFrameComparator:
    def __init__(self, df1, df2):
        self.df1 = df1
        self.df2 = df2
        self.df_diff = None

    def pre_sort(self, df, key_name):
        df.sort_values(by=key_name, inplace=True)

    def subract(self, only_key=True, key_col_name=''):

        if only_key:
            # keys = self.df1.merge(self.df2, how='inner', on=key_col_name)[key_col_name]

            self.pre_sort(self.df1, key_col_name)
            self.pre_sort(self.df2, key_col_name)
            # df_tmp = self.df1[self.df1[key_col_name].isin(keys)]
            df1_value = self.df1[self.df2.duplicated(key_col_name, keep=False)]
            df2_value = self.df2[self.df1.duplicated(key_col_name, keep=False)]
            # df_tmp =
            # df1_value = self.df1[self.df1[key_col_name].isin(keys)].select_dtypes(include='number')
            # df2_value = self.df2[self.df2[key_col_name].isin(keys)].select_dtypes(include='number')
        else:
            df_tmp = self.df1
            df1_value = self.df1.select_dtypes(include='number')
            df2_value = self.df2.select_dtypes(include='number')

        print(df1_value)
        print(df2_value)

        diff_value = df1_value - df2_value
        print(diff_value)
        for col in diff_value.columns:
            df_tmp[col] = diff_value[col]

        self.df_diff = df_tmp

if __name__  ==  '__main__':
    df1 = pd.DataFrame({'key' : ['aaa', 'bbb', 'ccc', 'ddd'],
            'A' : 1.,
                        'B' : pd.Timestamp('20130102'),
                        'C' : pd.Series(1,index=list(range(4)),dtype='float32'),
                        'D' : np.array([3] * 4,dtype='int32'),
                        'E' : pd.Categorical(["test","train","test","train"]),
                        'F' : 'foo' })
    df2 = pd.DataFrame({'key' : ['aaa', 'ccc', 'ggg', 'eee'],
            'A' : 1.,
                        'B' : pd.Timestamp('20130102'),
                        'C' : pd.Series(1,index=list(range(4)),dtype='float32'),
                        'D' : np.array([3] * 4,dtype='int32'),
                        'E' : pd.Categorical(["test","train","test","train"]),
                        'F' : 'foo' })

    # inst = DataFrameComparator(df1, df2)
    # inst.subract(False)
    # print(inst.df_diff)

    inst = DataFrameComparator(df1, df2)
    inst.subract(key_col_name='key')
    print(inst.df_diff)