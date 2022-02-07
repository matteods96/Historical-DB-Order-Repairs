# HistoricalDBOrderRepairs
AIM OF THE TASK:

Regarding this task I was suposed to create an historical db for order reparis related to compressor and turbines item.
In fact the data gathered since 2011 lack of information relevant to the client. In addition to this, there were a lot of incorrect or inaccurate data.
Thus, to  overcome this issue , the main idea was to combine the original data source related to the repairs with multiple data sources produced by ERP and tool application.
At the end, trought several combination this lead to create and hisotrical db made up of 3 different levels.
The first one (HEADER)  contains essential information referring to the repair order while in the second level there are all the details related to the repairs. 
Finally, the last one diplay whether or not  the SOGGE, which is the client adress, has been retrieved thanks to a certain join condition.
All in all, throughout retrieve a huge amount of information linked to the clients, the team was able to run a machine learning model to predict the customer profile as well to allow the development of further analysis related to the product future demand

(For further details read slide 1 to 9 of the projectTHESIS.pdf which might provide you a clearer overview of the task)


-FILES INSIDE THE FOLDER:
a) The .csv file and .pickle file are the main data sources which I have to analyze.
In particular at the beginning, I loaded  the excel file and then transformed into .pickle.
The  main reason relies on the fact that loading .pickle file  rather than excel files as the second type are quite heavy in terms of memory.
As a result in the scrip in Python I mainly loaded the .pickle files.

b)The repairs_aggreagations_matteo.py is the python script that enable to create the historical DB ( which is the Final_Output_RepairsFile.xlsx)




