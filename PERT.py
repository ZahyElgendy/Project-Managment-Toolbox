import re
import pandas as pd
import numpy as np

#Variable named in Camel Case format

#Define a class object named Task:
class Task(object):
    def __init__(self,description, activity, predecessors, duration):   #Initializes the attributes of task object
        #Attributes of a task
        self.description = description
        self.activity = activity.upper()
        self.predecessors = predecessors
        self.duration = duration
        self.earlyStart = 0
        self.earliestFinish = 0
        self.successors = []
        self.latestStart = 0
        self.latestFinish = 0
        self.slack = 0
        self.critical = ''

        #compute slack for a task
        # Slack is the diffrenrece between early finish and early start
        # or late finish and late start.

    #Task Method
    def computeSlack(self):
        self.slack = self.latestFinish - self.earliestFinish
        if self.slack > 0:     #if the slack is greater than zero then its non critical activity
            self.critical = 'NO'
        else:                  #otherwise it is critical
            self.critical = 'YES'


#function to get data from excel and return a pandas data frame:
def readData(excelFile, sheetname):
    mydata = pd.read_excel(excelFile, sheet_name= sheetname)
    return mydata

#function to compute duration the task or activity will be completed:
def computeDuration(mydata):
    mydata['DURATION'] = np.ceil((mydata['OPT'] + mydata['MOST']*4 + mydata['PESS'])/6)      #The duration is calculated by (optimistic + 4*most_likely + Pessimistic)/(6)
    return mydata

#function to create a task object (Instantiation):
def createTask(mydata): #recieves the dataframe
    taskObject =[]      #Creates empty list to contain the tasks

    for i in range(len(mydata)): #To iterate over each activity by iterating over each row i
        taskObject.append(Task(mydata['DESCRIPTION'][i], mydata['ACTIVITY'][i], mydata ['PREDECESSORS'][i], mydata['DURATION'][i]))
        # print(taskObject[i].description)
        # print(taskObject[i].activity)
        # print(taskObject[i].predecessors)
        # print(taskObject[i].duration)
    
    return (taskObject)

#function forward pass:
def forwardPass(taskObject):
    #Conjugate forward pass of cpm. Requires list of task objects
    #Rule 1: The early start of an activity with no predecssors is zero
    #Rule 2: The early start of an activity with predecessors is the maxiimum early Finish of its predecessors
    #Rule 3: The early finish of an activity is the sum of the early start and its activity time
    for task in taskObject:
        #if there an predecessor exits
        if type(task.predecessors) is str: #type string
            #make the string uppercase:
            task.predecessors = task.predecessors.upper()
            ef = [] # store the earliest finish of all the task's predecessors.
            #Get the maximum earlyFinish
            for j in task.predecessors: #for multiple predecessors
                for t in taskObject:                 # to iterate over all tasks
                    if t.activity == j:             
                        ef.append(t.earliestFinish)  #to append the earliest finish of all predecessors of current activity
                task.earlyStart = max(ef)            #to extract the maximum earliest finish date of all predecessors of current activity
            del ef

        else:   #else no predecessor exits, assign its early start to zero     
            task.earlyStart = 0                      

        # common for both situations:
        task.earliestFinish = task.earlyStart + task.duration         # the earliest finish can be computed by summing the start date and the duration of activity


#Function for backward pass:
def backwardPass(taskObject):
    #Implement backward pass, forwardPass should be implemented first
    #Returns the Latest Finsih and Latest Start time:
    #Rule 1: Latest Finish of the activity which is not a predecessor of any activity
    #is the maximum Earliest Finish.
    #Rule 2: Latest Finish of the activity which is a predecessor of any
    #activity is the minimum Earliest Finish of its successors
    #Rule 3: The Latest Start of the activity is Latest Finish less the Duration Time.

    pred = []
    eF = []

    #This For Loop will gather earliest Finish of all tasks and fill the predecessors list
    #This will also compute for the successors of each task : Task.successors variable
    for task in taskObject:
        if type(task.predecessors) == str:  #This means that a predecessor exits
            for j in task.predecessors: #for multiple predecessors
                pred.append(j) # fill in predecessor list of all tasks
                #fill in successors:
                for m in taskObject:
                    if m.activity ==j:
                        m.successors.append(task.activity)     #appends the successors of activity, since "successors" is a list in the task class as defined initially

        eF.append(task.earliestFinish)     # appends each task earliestFinish to ef list

    for task in reversed(taskObject): #Reverse iteration needed
        if task.activity not in pred:
            task.latestFinish = max(eF)    

        else:

            minLs = []
            for x in task.successors:
                for t in (taskObject):
                    if t.activity == x:
                        minLs.append(t.latestStart)

            task.latestFinish = min(minLs) #to extract the minimum late start date of all predecessors of current activity
            del minLs

        # common for both situtations:
        task.latestStart = task.latestFinish - task.duration       # the latest start can be computed by subtracting the task duration from its latest finish 

#function to compute for slack:
def slack(taskObject):
    for task in taskObject:
        task.computeSlack()   # a function in the task class

#update dataframe function:
def updateDataFrame(df, TaskObject): #Update data frame.
    df2 = pd.DataFrame({
        'DESCRIPTION' : df['DESCRIPTION'],
        'ACTIVITY' : df ['ACTIVITY'],
        'PREDECESSORS' : df ['PREDECESSORS'],
        'SUCCESSORS' : [task.successors for task in TaskObject],
        'OPT' : df ['OPT'],
        'MOST' : df ['MOST'],
        'PESS' : df ['PESS'],
        'DURATION' : df ['DURATION'],
        'ES' : pd.Series([task.earlyStart for task in TaskObject]),         #pd.Series converts One-dimensional array into coloumn in dataframe
        'EF' : pd.Series([task.earliestFinish for task in TaskObject]),
        'LS' : pd.Series([task.latestStart for task in TaskObject]),
        'LF' : pd.Series([task.latestFinish for task in TaskObject]),
        'SLACK' : pd.Series([task.slack for task in TaskObject]),
        'CRITICAL?' : pd.Series([task.critical for task in TaskObject]),
        'Cost$' : df ['Cost$'],
    })
    return(df2)


#function to compute for slack:




# function main:
def PERT(path):
    # get data and return a data frame df:
    df = readData(path, "Sheet1")
    print("Loaded Data: ")
    print(df)
    #Compute for the dats of the task:
    df = computeDuration(df)
    #Create task objects, Prior to this, create a Class named Tasks.
    taskObject = createTask(df)
    forwardPass(taskObject)
    backwardPass(taskObject)
    slack(taskObject)
    # Update data frame df to include new columns:
    finaldf = updateDataFrame(df, taskObject)
    print("\nResults:")

    #Pert Code inspired by PinoyStat

    #sort the tasks based on ES
    inOrder = finaldf.sort_values('ES')
    print(inOrder)


    #compute the total project duration based on critical activities
    criticalActivities = finaldf[finaldf["CRITICAL?"]=="YES"]
    #print(criticalActivities)
    Total = criticalActivities['DURATION'].sum()
    print(f"Total project duration is: {Total} unit time")



    criticalPath=criticalActivities.get('ACTIVITY').tolist()
    print(f"The critical path is : {criticalPath}")

    #save to file:
    print('\nResults saved to PERT.xlsx\n')
    finaldf.to_excel('PERT.xlsx', index = False)



                

