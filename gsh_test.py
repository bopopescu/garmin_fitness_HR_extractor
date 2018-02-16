import pygsheets

gc = pygsheets.authorize()

# Open spreadsheet and then workseet
sh = gc.open('FIT_DATA_TEST1')
wks_earning = sh.worksheet_by_title("Earning_history")

# Update a cell with value (just to let him know values is updated ;) )

activityID = 2496725900

heights = wks_earning.get_col(2)

if str(activityID) in heights:
    print(heights)
