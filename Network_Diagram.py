
import graphviz  #Main functions used node & edge
#node is used to specify nodes i.e., cirlces
#edge is used to connect between nodes i.e., arrows
import pandas as pd

def graphnetwork(path):
    mydata = pd.read_excel(io = path, sheet_name= 'Sheet1')
    dot = graphviz.Digraph('Network Diagram', comment='Network Diagram', graph_attr={'rankdir':'LR'})  
    # LR means that the graph will be oriented from left to right

    #This initializes all the nodes with their specific ACTIVITY/CODE names
    rows = len(mydata.axes[0]) #axes[0] means the rows while axes[1] means coloumns

    #zip combines elements from two or more lists
    for activity, number in zip(mydata['ACTIVITY'],  range(1, rows +1)):
        number = str(number)
        #Assigning each activity with a specific node number
        dot.node(number, activity)
        



    #This initializes "START" as node 0, and "END" as node -1
    dot.node('0', 'START')
    dot.node('-1', 'END')


                #The following code draws the arrows or edges between the nodes:
    #Adds edges to all Activites with no predecessor to be connected with 'START' (0)
    startingActivities = mydata[mydata['PREDECESSORS'].isnull()].index.tolist()  # makes a list of null predecessors i.e., empty 
    for i in startingActivities:
        #Assigns a number to the starting activities
        startingActivities_Num = str(i+1) 
        #makes an edge between "START" node 0 and activities with no predecessors
        dot.edge('0', startingActivities_Num)

    #Adds edges to all Activites with having successors
    for activity in mydata['ACTIVITY']:
        #converts activity into a string
        activity = str(activity)

        #Gets the successors for each activity
        successors=mydata[mydata['ACTIVITY']==activity]['SUCCESSORS'].values[0]
        
        #Gets the index of the current activity
        row_num = (mydata[mydata['ACTIVITY'] == activity]).index.tolist()
        #Since the "START" node is 0, all activities are numbered from 1
        activityIndex = str(row_num[0] + 1)

        #To convert the successors list to a string
        successors = str(successors)
        
        #To remove any non alphabetical characters (cleaning our data)
        #print(successors)
        char_remove = ["[", "]", ",", "'", " ",]
        for char in char_remove:
            successors = successors.replace(char, '')

        #Adds edges to all Activites that have no successors to be connected with "END"
        if len(successors) == 0: #this checks if successor value is NaN
            dot.edge(activityIndex,'-1')

            pass

        #Adds edges to all Activites that having successors
        else:
            for i in successors: #to iterate each successor and edges it with its predecessor.
                successorIndex = (mydata[mydata['ACTIVITY'] == i]).index.tolist()
                successorIndex = str(successorIndex[0] + 1)
                #print(successorIndex)
                dot.edge(activityIndex,successorIndex)

    #byadkhol 3la kol activity byshoof el successor wa ygeeb el index bt37a

    dot.render(directory='doctest-output_2', view=True)   #saves the network diagram to pdf
    'doctest-output/round-table.gv.pdf'





