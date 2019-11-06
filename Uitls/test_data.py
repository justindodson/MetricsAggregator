import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
df = pd.read_excel('C:/Users/204055604/Desktop/metrics/9.19.2919 Call/Training/TT_Report_2019-09-18.xlsx')

df = df.fillna(0)
df = df.drop([26, 27, 28], axis=0)

comp_array = np.array(df['Comp Rate'])
for i, val in enumerate(comp_array):
    comp_array[i] = val * 100

training_completions = pd.Series(data=comp_array, index=df['Dept'] )
labels = training_completions.keys().tolist()
rates = training_completions.values.tolist()

fig, bar = plt.subplots()
x = np.arange(len(rates))

rect = bar.bar(x, height=rates, data=rates, label='Completion Rate')

bar.set_ylabel('Completion Rates')
bar.set_title('Sout Region Training Completion')
plt.xticks(x, labels, rotation='vertical')

plt.rcParams.update({'font.size': 22})

bar.set_xticks(x)
bar.set_xticklabels(labels)
bar.legend()

plt.show()