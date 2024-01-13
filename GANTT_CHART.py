import plotly.express as px
import pandas as pd
                                                #File of PERT_CPM
def Gant(path):



                                                #Early start & Early Finish Schedule
    df = pd.read_excel(path)  
    tasks = df["DESCRIPTION"]                   #Create a list of all the tasks' description


    df['Duration'] = df['EF'] - df['ES']        #Calculates the duration which is the subtraction of  early "start" and  "finish"

    fig = px.bar(df, 
    color="CRITICAL?",      #Colours the bars based on if activities are critical or not
    base = "ES",            #Starts each bar from the early start
    x = "Duration",         #Sets the x to Duration
    y = tasks,              #Sets the Y-axis to tasks list                        
    orientation = 'h',      #Orients the bars horrizontally instead of vertically
    title="Gantt Chart for Early-Start Schedule",  #Titles the graph
    text = str("Duration")   #Annotate Bar
    )
    fig.update_traces(textfont_size = 24, textangle = 0, textposition = "auto", insidetextanchor = "middle")     #To modify the text font, angle, position and anchor
    fig.update_layout( yaxis={'categoryorder':'array', 'categoryarray':tasks})   #orders the y-axis according to the order of "tasks" array, otherwise could be set ascending or descending.
    fig.update_yaxes(autorange="reversed")    #This reverses the order of activities in Y coloumn
    fig.show()  #displays the bar chart



    #Late start & Late Finish Schedule
    df2 = pd.read_excel(path)
    tasks = df2["DESCRIPTION"]



    df2['Duration'] = df2['LF'] - df2['LS']

    fig2 = px.bar(df2,
    color="CRITICAL?",           
    base = "LS",
    x = "Duration",
    y = tasks,
    orientation = 'h',
    title="Gantt Chart for Late-Start Schedule",
    text = str("Duration")   #Annotate Bar
    )

    fig2.update_traces(textfont_size = 24, textangle = 0, textposition = "auto", insidetextanchor = "middle")     #To modify the text font, angle, position and anchor
    fig2.update_layout( yaxis={'categoryorder':'array', 'categoryarray':tasks})   #orders the y-axis according to the order of "tasks" array, otherwise could be set ascending or descending.
    fig2.update_yaxes(autorange="reversed")    #This reverses the order of activities in Y coloumn
    fig2.show()  #displays the bar chart



