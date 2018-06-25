##############################################################################
############################## File Description ##############################
##############################################################################
# Program to create pie graph of participation stress test for a union
# organizing campaign. The spreadsheet I used with this data stores the name
# of the employee in the first column (the name functions as the key for
# the other columns), the leader of the employee in the second column (store
# "null" if they do not have a leader, "other" if they are led by someone not
# in the spreadsheet, and "Unmapped" if they have not been assigned a leader
# yet). Other columns in the spreadsheet contain a 1 if the employee 
# participated in a specific stress test and 0 if they did not.
#
# This program is designed to display participation in the stress test 
# which has data stored in the third column of the spreadsheet.
#
# by: Liz Bishop
# Last modified: May 26, 2018
#
# Inspired by code from:
# - Python Graph Gallery 
#   (https://python-graph-gallery.com/163-donut-plot-with-subgroups/)
#
# Reusing code from:
# - Junwei's Lab
#   (http://junweihuang.info/uncategorized/
#    automatically-generate-n-distinct-colors-a-python-script/)
#       -> I used this code as a starting point for randomly assigning colors
##############################################################################

##############################################################################
# Import necessary libraries
##############################################################################
import matplotlib.pyplot as plt
import numpy as np
import xlrd
import sys

##############################################################################
# Initialize variables
##############################################################################
# The titles of the possible columns the user can select to graph
title_list = []
# The number identifying the column of the spreadsheet from which we are 
# pulling and graphing data
selected_col = 0
# Store names of all employees
name_list = []
# Store 1 if employee participated in the test, 0 if they did not
particp_list = []
# Store name of person who leads person at name_list[x]. If "null", this 
#   person leads themself
ledby_list = []
# A dictionary that maps each employee to their leader
Name_leader = {}
# A dictionary that maps each employee to 1 if they participated in the
#   stress test and 0 if they did not
Name_particip = {}
# A dictionary that maps leader names to the color associated with them
leader_color = {}
# A dictionary that maps each leader to a list of their followers
follower_map = {}
# A dictionary that maps each leader to the number of followers they have
follower_count = {}
# Store names of all employees who are not designated as leaders
followers = []
# the names of all leaders will be displayed in the inner ring
inner_ring = []
# color associated with the leader and all of their followers
inner_color = []
# Names in outer ring of pie chart - in order according to inner_ring
outer_ring = []
# color associated with each follower's leader (white indicates no 
# participation)
outer_color = []
# Store number of followers each leader has 
#   - number of followers inner_ring[0] has is stored at fol_count[0]
fol_count = []
# Number of columns in sheet - initialize to 0
col_length = 0

# Colors that we will assign to leaders
ColourValues=["FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", 
        "800000", "008000", "000080", "808000", "800080", "008080", "808080",
        "C00000", "00C000", "0000C0", "C0C000", "C000C0", "00C0C0", "C0C0C0",
        "400000", "004000", "000040", "404000", "400040", "004040", "404040",
        "200000", "002000", "000020", "202000", "200020", "002020", "202020",
        "600000", "006000", "000060", "606000", "600060", "006060", "606060",
        "A00000", "00A000", "0000A0", "A0A000", "A000A0", "00A0A0", "A0A0A0",
        "E00000", "00E000", "0000E0", "E0E000", "E000E0", "00E0E0", "E0E0E0"]
color_length=len(ColourValues)
white = "#ffffff"

##############################################################################
# Open relevate Excel spreadsheet, ask user which column they would like to
# see displayed on the graph, and gather title data.
##############################################################################
workbook = xlrd.open_workbook('Cafeteria_data.xlsx')
sheet = workbook.sheet_by_index(0)
col_length = len(sheet.col_values(0))
row_length = len(sheet.row_values(0))

# Ask user which column of data they would like to graph
index = 2
while index < row_length:
    title_list.append(sheet.cell(0, index))
    index += 1

user_choice = 0
print "What column of the spreadsheet would you like to graph?"

index = 0
while index < len(title_list):
    print "To display a pie chart containing data from spreadsheet column '" \
            + title_list[index].value + "', enter '" + str(index + 1) + \
            "' into your keyboard."
    index += 1

while not((selected_col <= row_length - 2) and selected_col >= 1):
    selected_col = input("Please enter a number between 1 and " + 
                        str(row_length - 2) + "!\n")

selected_col = selected_col + 1

header_data = sheet.cell(0, selected_col)
title_data = "Cafeteria Employees - " + header_data.value
##############################################################################

##############################################################################
# Transfer data from spreadsheet into lists and dictionaries.
##############################################################################
Names_data = sheet.col_slice(colx=0,
                              start_rowx=1,
                              end_rowx= col_length)

LedBy = sheet.col_slice(colx=1,
                              start_rowx=1,
                              end_rowx= col_length)

Partic_data = sheet.col_slice(colx=selected_col,
                              start_rowx=1,
                              end_rowx= col_length)

for cell in Names_data:
    name_list.append(cell.value)

for cell in LedBy:
    ledby_list.append(cell.value)

for cell in Partic_data:
    particp_list.append(cell.value)

Name_leader = dict(zip(name_list, ledby_list))
Name_particip = dict(zip(name_list, particp_list))

##############################################################################
# Determine whether each employee is a leader. If they are, assign them a 
# color. Create list of labels (leader names) and colors (the color associated
# with the leader if they participated, white  if they did not) to send to 
# graphing function to make inner circle of the pie chart. 
##############################################################################
# curr_row is a temporary variable used to assign colors to leaders
curr_row = 0
for cell in LedBy:
    if cell.value == "null":
        curr_color = "#"+ColourValues[np.mod(curr_row,color_length)]
        inner_ring.append(name_list[curr_row])
        leader_color[name_list[curr_row]] = curr_color
        if particp_list[curr_row] == 1:
            inner_color.append(curr_color)
        else:
            inner_color.append(white)
    else:
        followers.append(name_list[curr_row])
    curr_row += 1

# to account for fact that not all employees conform to conventional leadership
# hierarchy
inner_ring.append("Other")
inner_ring.append("Unmapped")
inner_color.append(white)
inner_color.append(white)
leader_color["Other"] = "#"+ColourValues[np.mod(curr_row,color_length)]
leader_color["Unmapped"] = "#"+ColourValues[np.mod(curr_row + 1,color_length)]

##############################################################################
# Create a list of follower names (in order according to leader) and colors 
# (the color of the leader of the employee if they particpated in the
# campaign, white if they did not) to send to the graphing function to make
# the outer circle of the pie chart
##############################################################################
for leader in inner_ring:
    follower_count[leader] = 0
    follower_map[leader] = []

for follower in followers:
    follower_count[Name_leader[follower]] += 1
    follower_map[Name_leader[follower]].append(follower)

for leader in inner_ring:
    for follower in follower_map[leader]:
        outer_ring.append(follower)
        if Name_particip[follower] == 1:
            outer_color.append(leader_color[leader])
        else:
            outer_color.append(white)
    fol_count.append(follower_count[leader])

##############################################################################
# Graph the two pie charts
##############################################################################
# First Ring (outside)
fig, ax = plt.subplots()
ax.axis('equal')
mypie, _ = ax.pie(np.full((len(outer_ring)), 1), radius=1.3, 
                                labels=outer_ring, colors=outer_color)
plt.setp( mypie, width=0.3, edgecolor='gray')

# Second Ring (Inside)
mypie2, texts = ax.pie(fol_count, radius=1.3-0.3, labels=inner_ring, 
                                labeldistance=0.75, colors=inner_color)
for t in texts:
    t.set_horizontalalignment('center')
plt.setp( mypie2, width=0.4, edgecolor='gray')
plt.margins(0,0)
plt.title(title_data, y=1.08)

# Display the graph
plt.show()