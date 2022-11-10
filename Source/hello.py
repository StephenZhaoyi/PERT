#Zhao Yi 2009853G-I011-0067
#import libraries
import pandas as pd
import numpy as np

#define a class, each class represents each activity

class Act(object):
    def __init__(self,description,activity,predecessors,esti):
        self.descirption = description
        self.activity = activity
        self.predecessors = predecessors
        self.esti = esti

        self.earliestStart = 0
        self.earliestFinish = 0
        self.successors = []
        self.latestStart = 0
        self.latestFinish = 0
        self.slack = 0
        self.critical = ''
#compute slack
    def computeSlack(self):
        self.slack = self.latestFinish - self.earliestFinish
        if(self.slack == 0):
            self.critical = 'Y'
        else:
            self.critical = 'N'

    

#read date from input.xlsx using "pandas"

mydata = pd.read_excel("Input.xlsx")

#compute estimated time, choose the upper bound
def computeEsti(data):
    data['Estimated'] = np.ceil((data['Optimistic'] + 4*data["Most likely"] + data["Pessimistic"])/6)
    return data

#create action object
def createAct(data):
    actObject = []
    for i in range(len(data)):
        actObject.append(Act(data['Description'][i],data['Activity'][i],data["Predecessors"][i],data['Estimated'][i]))
    return actObject

#initialize and compute
def initialize(data,act):
    for i in range(len(data)):
        if (isinstance(act[i].predecessors,float)):
            act[i].earliestStart = 0
            act[i].earliestFinish = act[i].esti
        else:
            ef = []
            for j in act[i].predecessors:
                for k in act:
                    if (j == k.activity):
                        ef.append(k.earliestFinish)
                        k.successors.append(act[i].activity)
            act[i].earliestStart = max(ef)
            act[i].earliestFinish = (act[i].earliestStart + act[i].esti)

    count = len(data)
    for i in range(count-1,-1,-1):
        if(len(act[i].successors) == 0):
            act[i].latestFinish = act[count-1].earliestFinish
            act[i].latestStart = act[i].latestFinish - act[i].esti
        else:
            lf = []
            for j in (act[i].successors):
                for k in (act):
                    if(j == k.activity):
                        lf.append(k.latestStart)
            act[i].latestFinish = min(lf)
            act[i].latestStart = act[i].latestFinish - act[i].esti
 
    return act

#output

def output(data,act):
    for i in range(len(data)):
        data['ES'][i] = act[i].earliestStart
        data['EF'][i] = act[i].earliestFinish
        data['LS'][i] = act[i].latestStart
        data['LF'][i] = act[i].latestFinish
        act[i].computeSlack()
        data['Slack'][i] = act[i].slack
        data['Critical?'][i] = act[i].critical
    return data
print("Before")
print(mydata)
#core
#compute estimated time
mydata = computeEsti(mydata)

#set action
action = createAct(mydata)

#initialize
action = initialize(mydata,action)

#output
mydata = output(mydata,action)
mydata.to_excel("output.xlsx")

print("After:")
print(mydata)

