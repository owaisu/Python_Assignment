    with open(os.path.join(path,"Part1/CSV file output/Task1.csv"),'w') as csf:
        csw=csv.writer(csf,delimiter='|')
        csw.writerow(head)
        # csw.writerows(rows)
except UnicodeError:
    pass