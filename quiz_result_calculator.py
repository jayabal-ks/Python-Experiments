import pandas as pd
import numpy as np

dataframe = pd.read_excel('DevOps Challenge _ Quiz -1(1-19).xlsx')

resultdf = pd.DataFrame(columns=['Name','EmpId','Time','Points','Q1','Q2','Q3','Q4','Q5','Q6','Q7','Q8','Q9','Q10','Q11','Q12','Q13','Q14','Q15'])

columns = dataframe.columns

points_columns = []

for i in range(columns.size):
   if columns[i].startswith("Points - "):
    points_columns.append(columns[i-1])
    points_columns.append(columns[i])

points_columns = points_columns[4:]

for index, row in dataframe.iterrows():
    name = row["Name"]
    EmpId = row["Employee ID"]
    total_time = row["Completion time"] - row["Start time"]
    points = []

    for i in range(1, len(points_columns), 2):
       if (row[points_columns[i]] == 1):
        points.append(1)
       elif (row[points_columns[i]] == 0 and row[points_columns[i-1]].strip() == ""):
         points.append(0)
       else:
         points.append(-0.5)
    
    df2 = {'Name': name, 'EmpId':EmpId, 'Time': total_time.total_seconds()/60, 'Points': sum(points),
            'Q1':points[0],
            'Q2':points[1],
            'Q3':points[2],
            'Q4':points[3],
            'Q5':points[4],
            'Q6':points[5],
            'Q7':points[6],
            'Q8':points[7],
            'Q9':points[8],
            'Q10':points[9],
            'Q11':points[10],
            'Q12':points[11],
            'Q13':points[12],
            'Q14':points[13],
            'Q15':points[14],
            }

    resultdf.loc[len(resultdf)] = df2
  
print(resultdf)

resultdf.to_excel('result.xlsx')