import pandas as pd
import os

#Remember to put the script in the same directory as .xlsx files!
directory = os.getcwd()

#Empty list for uses
locations = []
num_people = []
positions = []

#Read names
for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):
        #Parse the location, person, and position from the filename
        file_parts = filename.split("_")
        location = file_parts[0] + " "
        num = 1
        position = file_parts[-1].split(".")[0] #index [-1] is the last element of the list

        #Append the location, number of people, and position to list
        locations.append(location)
        num_people.append(num)
        positions.append(position)

#Combine three lists into a DataFrame and rearrange the columns
summary_df = pd.DataFrame({"Location": locations, "Number of people": num_people, "Position": positions})

#Group the DataFrame by location and position and sum the number of people (this is to get the number of people!)
summary_df = summary_df.groupby(["Location", "Position"]).sum()
#if the line below is not here = the list of index will be error after sum
summary_df = summary_df.groupby(["Location", "Position"]).sum().reset_index()
#Rearrange the columns to move "Number of people" to the middle
summary_df = summary_df[["Location", "Number of people", "Position"]]

# Write the summary table to a text file
with open("summary.txt", "w") as f:
    f.write(summary_df.to_string(index=False))

# Print the summary table
print(summary_df.to_string(index=False))

