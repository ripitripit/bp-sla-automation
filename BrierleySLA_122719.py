import pandas as pd
import numpy as np
import datetime

pd.read_excel('SD_ALL_SLA.xlsx', 'Sheet1', index_col='Created', usecols=['Created','Reporter','Assignee','Custom field (Time to first response)','Custom field (Time to resolution)']).to_csv(
    'SD_ALL_SLA.csv', encoding='utf-8')

df = pd.read_csv('SD_ALL_SLA.csv', delimiter=',', usecols=['Created','Reporter','Assignee','Custom field (Time to first response)','Custom field (Time to resolution)'], encoding='utf-8')

df['month'] = pd.DatetimeIndex(df['Created']).month
df['year'] = pd.DatetimeIndex(df['Created']).year 

nearshore = ['ajarrell', 'bpatrick', 'aknapp', 'ctippet', 'jgaribello', 'jway', 'jweiner', 'kpatrick', 'mvandelagemaat', 'rbedlack', 'rsirigiri', 'spatel', 'uperrotta', 'dpatel', 'TampaSD']

nonNearshore=[]
list1 = df['Reporter'].tolist()
list2 = df['Assignee'].tolist()
list3 = list1 + list2
for x in list3:
    if x not in nonNearshore and x not in nearshore:
        (nonNearshore.append(x))
nonNearshore = [str(x) for x in nonNearshore]

year = int(input('Enter the year to report on. IE <2019>: ' ))
month = int(input('Enter the month to report on. IE <6>: ' ))

indexNames = df[df['year'] != year].index
indexNames2 = df[df['month'] != month].index
df.drop(indexNames, inplace=True)
df.drop(indexNames2, inplace=True)

indexNames3 = df[df['Custom field (Time to first response)'] == 'Request Type = No Match'].index
df.drop(indexNames3, inplace=True)

def findCt(column1, list1):
    count = 0
    for name in list1:
        count += np.sum(df[column1].str.contains(name))
    return count

def findCt2(column1, column2, list1):
    count = 0
    for name in list1:
        count += np.sum((df[column1].str.contains(name, na=False) & (~df[column2].str.contains('-', na=True))))
    return count

nearRepCt = findCt('Reporter', nearshore)
nearRepResponseCt = findCt2('Reporter', 'Custom field (Time to first response)', nearshore)
nearRepResolveCt = findCt2('Reporter', 'Custom field (Time to resolution)', nearshore)
nearRepResponseSLA = '{:.2f}'.format((nearRepResponseCt / nearRepCt) * 100)
nearRepResolveSLA = '{:.2f}'.format((nearRepResolveCt / nearRepCt) * 100)

nearAssCt = findCt('Assignee', nearshore)
nearAssResponseCt = findCt2('Assignee', 'Custom field (Time to first response)', nearshore)
nearAssResolveCt = findCt2('Assignee', 'Custom field (Time to resolution)', nearshore)
nearAssResponseSLA = '{:.2f}'.format(nearAssResponseCt / nearAssCt * 100)
nearAssResolveSLA = '{:.2f}'.format(nearAssResolveCt / nearAssCt * 100)

print('')
print('Nearhshore by Reporter', ('(' + '0'  + str(month) + '-' + str(year) + ')'))
print('Total# of tickets created:', nearRepCt)
print('Tickets Met for Time to First Response:', nearRepResponseCt)
print('Tickets Met for Time to Resolve:', nearRepResolveCt)
print('SLA% for Time to First Response:', nearRepResponseSLA)
print('SLA% for Time to Time to Resolve:', nearRepResolveSLA)
print('')
print('Nearhshore by Assignee', ('(' + '0'  + str(month) + '-' + str(year) + ')'))
print('Total# of tickets created:', nearAssCt)
print('Tickets Met for Time to First Response:', nearAssResponseCt)
print('Tickets Met for Time to Resolve:', nearAssResolveCt)
print('SLA% for Time to First Response:', nearAssResponseSLA)
print('SLA% for Time to Time to Resolve:', nearAssResolveSLA)

nonRepCt = findCt('Reporter', nonNearshore)
nonRepResponseCt = findCt2('Reporter', 'Custom field (Time to first response)', nonNearshore)
nonRepResolveCt = findCt2('Reporter', 'Custom field (Time to resolution)', nonNearshore)
nonRepResponseSLA = '{:.2f}'.format((nonRepResponseCt / nonRepCt) * 100)
nonRepResolveSLA = '{:.2f}'.format((nonRepResolveCt / nonRepCt) * 100)

nonAssCt = findCt('Assignee', nonNearshore)
nonAssResponseCt = findCt2('Assignee', 'Custom field (Time to first response)', nonNearshore)
nonAssResolveCt = findCt2('Assignee', 'Custom field (Time to resolution)', nonNearshore)
nonAssResponseSLA = '{:.2f}'.format(nonAssResponseCt / nonAssCt * 100)
nonAssResolveSLA = '{:.2f}'.format(nonAssResolveCt / nonAssCt * 100)

print('')
print('non-Nearshore by Reporter', ('(' + '0'  + str(month) + '-' + str(year) + ')'))
print('Total# of tickets created:', nonRepCt)
print('Tickets Met for Time to First Response:', nonRepResponseCt)
print('Tickets Met for Time to Resolve:', nonRepResolveCt)
print('SLA% for Time to First Response:', nonRepResponseSLA)
print('SLA% for Time to Time to Resolve:', nonRepResolveSLA)
print('')
print('non-Nearshore by Assignee', ('(' + '0'  + str(month) + '-' + str(year) + ')'))
print('Total# of tickets created:', nonAssCt)
print('Tickets Met for Time to First Response:', nonAssResponseCt)
print('Tickets Met for Time to Resolve:', nonAssResolveCt)
print('SLA% for Time to First Response:', nonAssResponseSLA)
print('SLA% for Time to Time to Resolve:', nonAssResolveSLA)
