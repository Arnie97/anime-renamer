dir /b >rename.csv
start /wait rename.csv

REM Copy the column in the list created.
REM Manually replace filenames in the new column.
REM Create a column filled with "ren ".

ren rename.csv rename.cmd
start /wait notepad rename.cmd

REM Replace 'ren ,' with 'ren "', ',\r\n' with '"\r\n' and ','' with '" "'.

rename.cmd
