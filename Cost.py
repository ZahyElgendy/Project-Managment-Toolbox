import plotly.express as px
import pandas as pd
import plotly.graph_objects as go


def cost(path):

    finaldf = pd.read_excel(path)
    criticalActivities = finaldf[finaldf["CRITICAL?"]=="YES"]    #To get the critical activities only
    Total = criticalActivities['DURATION'].sum()                 #To get the total duration of the project by summing the durations of each critical activity


#Early Start Cost Analysis


#Creating a new Dataframe consisting of Duration as coloumn 
    coloumns = ['DURATION']
    for i in finaldf['ACTIVITY']:
        coloumns.append(i)
      

    costsES_df = pd.DataFrame(columns=coloumns)
    #print(costsES_df)
    #Creating a list with the project's DURATION by each unit of time
    project_DURATION = list(range(1, int(Total)+1)) # +1 since range counts a number before the last integer
    #Assigning the list to 'DURATION' column
    costsES_df.loc[:, 'DURATION'] = project_DURATION

    #To fill with each activity cost coresponding to each day
    for i_cost in range(len(finaldf)):
        #assign each day from the Early start (ES) to the Early Finish (EF) with the cost of the activity
        #Cost of the activity for each day is its total cost divided by the duration of the activity. For examply  1200$/6days= 200$ each day
        costsES_df.iloc[int(finaldf['ES'][i_cost]):int(finaldf['EF'][i_cost]), i_cost+1] =  round(finaldf['Cost$'][i_cost]/finaldf['DURATION'][i_cost],3)

    #print(costsES_df)

    costsES_df= costsES_df.fillna(0) #to replace all null values in the costs_dataframe with zeros

    costsES_df['Daily Cost $ (ES)'] = '' #to create new coloumn for the Daily Cost
    costsES_df['Cumulative Cost $ (ES)'] = '' #to create new coloumn for the Cumulative Cost
    #print(costsES_df)
    
    coloumns.pop(0) #to remove the Duration Coloumn of unit time.

    #Computing Daily Cost
    costsES_df['Daily Cost $ (ES)'] = costsES_df.loc['0':Total,coloumns].sum(axis = 1)
    #df['New Coloumn'] = df.loc['initial row':'final row',[coloumns to sum]].sum(axis = 1)      #axis = 1 means col




    #Computing Cumulative Cost
    sum = 0
    for i_cum_cost in range(0, int(Total)): #to iterate over all rows
        #summing all previous cells of daily cost up to the ith day
        sum = sum + round(costsES_df['Daily Cost $ (ES)'][i_cum_cost],3)   #rounding the cost to the nearest 3rd place

        #Assigning the sum to cumulative cost coloumn  corresponding to ith day
        costsES_df.iloc[i_cum_cost, -1] =  sum                              # -1 indicates the last coloumn which is "Cumulative Cost $ (ES)""
    
    print(f"Total Cost is {sum}")
    #print("The Daily & Cumulative Cost Computed (ES)")
    #print(costsES_df)




#Late Start Cost Analysis
    

#Creating a new Dataframe consisting of Duration as coloumn 
    coloumns = ['DURATION']
    for i in finaldf['ACTIVITY']:
        coloumns.append(i)
    costsLS_df = pd.DataFrame(columns=coloumns) #Assigning the list coloumns as the coloumns of costsLS_dataframe
    
    #Creating a list with the project's DURATION by each unit of time
    project_DURATION = list(range(1, int(Total)+1)) # +1 since range counts a number before the last integer
    #Assigning the list to 'DURATION' column
    costsLS_df.loc[:, 'DURATION'] = project_DURATION

    #To fill with each activity cost coresponding to each day
    for i_cost in range(len(finaldf)):
        #assign each day from the Early start (ES) to the Early Finish (EF) with the cost of the activity
        #Cost of the activity for each day is its total cost divided by the duration of the activity. For examply  1200$/6days= 200$ each day
        costsLS_df.iloc[int(finaldf['LS'][i_cost]):int(finaldf['LF'][i_cost]), i_cost+1] =  round(finaldf['Cost$'][i_cost]/finaldf['DURATION'][i_cost],3)

    #print(costsLS_df)

    costsLS_df= costsLS_df.fillna(0) #to replace all null values in the costs_dataframe with zeros

    costsLS_df['Daily Cost $ (LS)'] = '' #to create new coloumn for the Daily Cost
    costsLS_df['Cumulative Cost $ (LS)'] = '' #to create new coloumn for the Cumulative Cost
    #print(costsLS_df)
    coloumns.pop(0) #to remove the Duration Coloumn of unit time.

    #Computing Daily Cost
    costsLS_df['Daily Cost $ (LS)'] = costsLS_df.loc['0':Total,coloumns].sum(axis = 1)
    #df['Sum_Col'] = df.loc['r_i':'r_f',[col_to_sum]].sum(axis = 1)      #axis = 1 means col




    #Computing Cumulative Cost
    sum = 0
    for i_cum_cost in range(0, int(Total)):                               #to iterate over all rows
        sum = sum + round(costsLS_df['Daily Cost $ (LS)'][i_cum_cost],3)  #summing all previous cells of daily cost up to the ith day

        costsLS_df.iloc[i_cum_cost, -1] =  sum                       #Assigning the sum to cumulative cost corresponding to ith day
        
    #print("The Daily & Cumulative Cost Computed (LS)")
    #print(costsLS_df)



    #Saving the Early Start Costs and Late Start Costs to Excel file with different sheets
    with pd.ExcelWriter('Cost_Analysis.xlsx') as excel_writer:
        costsES_df.to_excel(excel_writer, sheet_name='Costs(ES)', index=False)
        costsLS_df.to_excel(excel_writer, sheet_name='Costs(LS)', index=False)





def generate_Graphs():
#Generating Graphs between Early Start & Late Start Cost Analysis



    costsES_df = pd.read_excel('Cost_Analysis.xlsx', sheet_name = 0)  #make dataframe for Early Start Costs
    costsLS_df = pd.read_excel('Cost_Analysis.xlsx', sheet_name = 1)  #make dataframe for Late Start Costs

                                    # Create Line plot for Daily Costs
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=costsES_df['DURATION'],   #Assign the X-axis to 'Duration' coloumn from the Early Start costs dataframe
        y=costsES_df['Daily Cost $ (ES)'], #Assign the Y-axis to 'Daily Cost $ (ES)' coloumn from the Early Start costs dataframe
        name="Daily Cost $ (ES)"    #Name the plot line as Daily Cost $ (ES)
    ))

    fig.add_trace(go.Scatter(
        x=costsLS_df['DURATION'],    #Assign the X-axis to 'Duration' coloumn from the Late Start costs dataframe
        y=costsLS_df['Daily Cost $ (LS)'],    #Assign the Y-axis to 'Daily Cost $ (LS)' coloumn from the Late Start costs dataframe  
        name="Daily Cost $ (LS)",    #Name the plot line as Daily Cost $ (LS)
        
    ))
    fig.update_layout(
        title=dict(text="ES vs. LS : Daily Cost $ ", font=dict(size=50), yref='paper')   #Put a title, specify the font size and type
    )

    # Display the plot
    fig.show()



    # Create Line plot for Cumulative Costs
    fig_2 = go.Figure()
    fig_2.add_trace(go.Scatter(
        x=costsES_df['DURATION'],  #Assign the X-axis to 'Duration' coloumn from the Early Start costs dataframe
        y=costsES_df['Cumulative Cost $ (ES)'], #Assign the Y-axis to 'Cumulative Cost $ (ES)' coloumn from the Early Start costs dataframe
        name="Cumulative Cost $ (ES)"  #Name the plot line as Cumulative Cost $ (ES)
    ))

    fig_2.add_trace(go.Scatter(
        x=costsLS_df['DURATION'],  #Assign the X-axis to 'Duration' coloumn from the Late Start costs dataframe
        y=costsLS_df['Cumulative Cost $ (LS)'],  #Assign the Y-axis to 'Cumulative Cost $ (LS)' coloumn from the Late Start costs dataframe  
        name="Cumulative Cost $ (LS)",    #Name the plot line as Cumulative Cost $ (LS)
        
    ))
    fig_2.update_layout(
        title=dict(text="ES vs. LS : Cumulative Cost $", font=dict(size=50), yref='paper')  #Put a title, specify the font size and type
    )

    # Display the plot
    fig_2.show()



