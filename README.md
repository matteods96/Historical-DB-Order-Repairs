# HistoricalDBOrderRepairs
Regarding this task I was suposed to create an historical db for order reparis related to compressor and turbines item.
In fact the data gathered since 2011 lack of information relevant to the client. In addition to this, there were a lot of incorrect or inaccurate data.
Thus, to  overcome this issue , the main idea was to combine the original data source related to the repair with multiple data sources produced by ERP and tool application.
At the end trought several combination this lead to create and hisotrical db made up of 3 different levels.
The first one (HEADER)  contains essential information referring to the repair order while in the second level there are all the details related to the repairs. 
Finally, the last one diplay when the SOGGE, which is the client adress, has been retrieved thanks to a certain join condition.
