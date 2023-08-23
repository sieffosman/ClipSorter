import pandas as pd

# Load the data from 'Revision_Lists_Year_9.xlsx' and 'MathsWatch_Clips.xls'
df1 = pd.read_excel('C:\\Users\\sieff\\Downloads\\clipsorter\\Revision_Lists_Year_9.xlsx')
df2 = pd.read_excel('C:\\Users\\sieff\\Downloads\\clipsorter\\MathsWatch_Clips.xls')


# Iterate over each row in df1
for index, row in df1.iterrows():
    # Get the Title from df1
    title = row['Title']
    
    # Find the matching row(s) in df2
    matches = df2[df2['Title'].str.contains(title, na=False, regex=False)]
    
    # Check if any matches are found
    if not matches.empty:
        # Get the first match
        match_row = matches.iloc[0]
        
        # Get the Clip Number from df2
        clip_number = match_row['Clip Number']
        
        # Update the Clip Number in df1
        df1.at[index, 'Clip Number'] = clip_number

# Print the updated dataframe in the console
print(df1)

# Save the updated dataframe to a new file
df1.to_excel('Updated_Revision_Lists_Year_9.xlsx', index=False)