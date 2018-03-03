# google-sheets
processing data in google sheets


### Requirements

1. read google sheetId and credentials from config file
2. open sheet
3. verify sheet has Date and Visitors columns
4. verify Moving Average column exists. if not, create.
5. for each row
Calculate moving average of visitors. The formula for moving average
is (Current Total Visitors / Number of Days). So for each row
encountered the Current Total Visitors will increase by the
new daily amount and the Number of Days will increase by 1 giving us
the new moving average.

### Possible Error Conditions
1. no data in Date field
2. no data in Visitors field
3. no Date field
4. no Visitors field
5. Visitors data type is not numeric
6. sheetId does not exist
7. sheetId is not readable
8. sheetId is not writable

