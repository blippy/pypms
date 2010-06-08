rem make database copy
set 7z="C:\Program Files\7-Zip\7z.exe"
set zfile=c:\mcarter\finance.7z
del %zfile%
7z a %zfile% @listfile.txt
