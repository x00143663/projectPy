import csv

sourcePath='/Users/naanast/Desktop/Project/task3/data/datasource.csv'

destPath='/Users/naanast/Desktop/Project/task3/data/ResultQ+.csv'

compPath='/Users/naanast/Desktop/Project/task3/data/qplus-quality-report.csv'




with open(destPath, mode='w') as result:
    result=csv.writer(result,delimiter=',',quotechar='"')
    result.writerow(['Building', 'Manager', 'Case_ID','Ratings','Reviewed','Profile','Amount'])

    with open(sourcePath) as source:
        readCSV = csv.reader(source, delimiter=',')
    

    with open(sourcePath,encoding='utf-16') as source:
        readCSV = csv.reader(source,delimiter='\t')
        # This skips the first row of the CSV file.

        next(readCSV)
        for row in readCSV:
           # print(row)


            building=row[6]
            manager=row[30]
            caseid=row[17]
            ratings=row[4]
            profile=row[35]
            count=0
            with open(compPath) as comp:
                compCSV = csv.reader(comp, delimiter=',')
                for row in compCSV:
                     if caseid==row[5]:
                        count=count+1

            if count != 0:
                reviewed = 1
            else:
                reviewed = 0

            result.writerow([building, manager, caseid, ratings, reviewed, profile, count])







