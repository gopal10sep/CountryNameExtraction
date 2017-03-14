import xlrd
import numpy as np
import matplotlib.pyplot as plt
import re
import random
import pandas as pd
import geograpy
from geograpy import extraction
#from geopy.geocoders import Nominatim
import pycountry

# Open the workbook
xl_workbook = xlrd.open_workbook("unmanned-outbound-logistics.xlsx") 

# Grabbing the Required Sheet by index 
xl_sheet = xl_workbook.sheet_by_index(0)

#Initialising the Lists Required

# topics = []
# currTopicList = []
# overallTopicList = []
# geolocator = Nominatim()
new1 = 'No Value'
new2 = ''
new3 = ''
new4 = ''
final ='No Value'
total = []
k=0
j=0



num_cols = xl_sheet.ncols   #Number of columns
for row_idx in range(2, xl_sheet.nrows):    # Iterate through rows
    #print (row_idx)
    country = []
    country2 = []
    country3 = []

    for col_idx in range(3, 4):  # Iterate through article_citation column
        cell_obj = xl_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        text2 = str(cell_obj)[7:-1]
    text2 = ''.join([i for i in text2 if not i.isdigit()])
    #print (text2)
    country2 = text2.split()
    country3 = [x.strip() for x in text2.split(',')]
    #print (country3)

 
    e = extraction.Extractor(text=text2)
    e.find_entities()
    #print ('******************')
    # You can now access all of the places found by the Extractor
    #print e.places
    #print (len(e.places))
    for i in range(0, len(e.places)):
        # print (e.places[i])

        try:
            new1 = pycountry.countries.get(alpha_2=e.places[i])
            final = new1.name
        except KeyError:
            try:
                new1 = pycountry.countries.get(alpha_3=e.places[i])
                final = new1.name
            except KeyError:
                try:
                    new1 = pycountry.countries.get(name=e.places[i])
                    final = new1.name
                except KeyError:
                    try:
                        new1 = pycountry.countries.get(official_name=e.places[i])
                        final = new1.name
                    except KeyError:
                        #final = 'No Value'
                        try:
                            new1 = pycountry.countries.get(common_name=e.places[i])
                            final = new1.name
                        except KeyError:
                            final = 'No Value'


        # print (final)
        # print ('##################')
        if (final != 'No Value' ):
            country.append(final)

    for i in range(0, len(country2)):
        # print (e.places[i])

        # if (country2[i] == 'Taiwan'):
        #     country2[i] = 'Taiwan, Province Of China'


        try:
            new1 = pycountry.countries.get(alpha_2=country2[i])
            final = new1.name
        except KeyError:
            try:
                new1 = pycountry.countries.get(alpha_3=country2[i])
                final = new1.name
            except KeyError:
                try:
                    new1 = pycountry.countries.get(name=country2[i])
                    final = new1.name
                except KeyError:
                    try:
                        new1 = pycountry.countries.get(official_name=country2[i])
                        final = new1.name
                    except KeyError:
                        #final = 'No Value'
                        try:
                            new1 = pycountry.countries.get(common_name=country2[i])
                            final = new1.name
                        except KeyError:
                            final = 'No Value'

        # print (final)
        # print ('##################')
        if (final != 'No Value' ):
            country.append(final)

    for i in range(0, len(country3)):
    # print (e.places[i])

    # if (country2[i] == 'Taiwan'):
    #     country2[i] = 'Taiwan, Province Of China'


        try:
            new1 = pycountry.countries.get(alpha_2=country3[i])
            final = new1.name
        except KeyError:
            try:
                new1 = pycountry.countries.get(alpha_3=country3[i])
                final = new1.name
            except KeyError:
                try:
                    new1 = pycountry.countries.get(name=country3[i])
                    final = new1.name
                except KeyError:
                    try:
                        new1 = pycountry.countries.get(official_name=country3[i])
                        final = new1.name
                    except KeyError:
        #final = 'No Value'
                        try:
                            new1 = pycountry.countries.get(common_name=country3[i])
                            final = new1.name
                        except KeyError:
                            final = 'No Value'

        # print (final)
        # print ('##################')
        if (final != 'No Value' ):
            country.append(final)


    #print (country)

    if (len(country) == 0):
        k = k+1
        country.append('No Value')
        #country[0] = 'A'
    else:
        j=j+1

    z= len(country)-1
    if (z<0):
        z=0

    my_list = [str(country[x]) for x in range(len(country))]




    total.append(my_list[z])

    #print (z)
#print (total)

# dictTopicList = {x:total.count(x) for x in total}
# import operator
# sorted_x = sorted(dictTopicList.items(), key=operator.itemgetter(1),reverse=True)

dictTopicList = {x:total.count(x) for x in total}
#print (dictTopicList)
value_to_remove = 'No Value'
dictTopicList =  {key: value for key, value in dictTopicList.items() if key != value_to_remove}
#print (dictTopicList)
import operator
sorted_x = sorted(dictTopicList.items(), key=operator.itemgetter(1),reverse=True)


i = 0
freq=[]
terms=[]
while ( i < len(sorted_x)):
    freq.append(float(sorted_x[i][1]))
    terms.append(sorted_x[i][0])
    # print (type(sorted_x[i][0]))
    # print (type(sorted_x[i][1]))
    i = i +1

#Plotting the Graph
N = len(sorted_x)
ind = np.arange(N)      #The x locations for the groups
width = 0.35            #The width of the bars
fig, ax = plt.subplots()
rects1 = ax.bar(ind, freq, width, color='r')
#rects2 = ax.bar(ind + width, patent_citation_count, width, color='y')

# #Add some text for labels, title and axes ticks

ax.set_ylabel('Frequency')
ax.set_title('Number of authors from different nations')
ax.set_xticks(ind)
ax.set_xticklabels(terms, rotation='vertical')
#ax.legend((rects1), ('Topic Modelled'))

#Attach a text label above each bar displaying its height
def autolabel(rects):    
    for rect in rects:
        height = rect.get_height()
        if height != 0 :
            ax.text(rect.get_x() + rect.get_width()/2., 1.05*height,'%d' % int(height),ha='center', va='bottom')
autolabel(rects1)


# Tweak spacing to prevent clipping of tick-labels
plt.subplots_adjust(bottom=0.50)
axes = plt.gca()
axes.set_ylim([0,np.max(freq)+10])
plt.savefig('histogram.png', dpi = 1000)
plt.show()





# #df = pd.DataFrame(list(sorted_x.iteritems()), columns=['terms','freq'])
# df = pd.DataFrame(sorted_x,columns=['terms','freq'])
# #Create a Pandas Excel writer using XlsxWriter as the engine.
# excel_file = 'histo.xlsx'
# sheet_name = 'Sheet1'

# writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
# df.to_excel(writer, sheet_name=sheet_name)

# # Access the XlsxWriter workbook and worksheet objects from the dataframe.
# workbook = writer.book
# worksheet = writer.sheets[sheet_name]

# # Create a chart object.
# chart = workbook.add_chart({'type': 'column'})

# # Configure the series of the chart from the dataframe data.
# chart.add_series({
#     'categories': ['Sheet1', 1, 1, 243, 1],
#     'values':     ['Sheet1', 1, 2, 243, 2],
  
# })


# # Configure the chart axes.
# chart.set_x_axis({'name': 'Country'})
# chart.set_y_axis({'name': 'Freq', 'major_gridlines': {'visible': False}})

# # Turn off chart legend. It is on by default in Excel.
# chart.set_legend({'position': 'none'})

# # Insert the chart into the worksheet.
# worksheet.insert_chart('G2', chart)

# #Close the Pandas Excel writer and output the Excel file.
# writer.save()