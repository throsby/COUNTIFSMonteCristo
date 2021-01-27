import pandas as pd
import os

start = "4"
end = "262"
york = "York"
wash = "Washington"
set = "Set"
locations = [york,wash,set]
# print('=COUNTIFS($C$%s:$C$%s, "<>y", $K$%s:$%s, "%s")'% (start,end,start,end,set))

d = {   'col0':["","Set","Washington","York"],
        'col1':[    "PCRs Left",
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR")'.format(start=start,end=end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR")'.format(start=start,end=end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR")'.format(start=start,end=end,location=york)],
        'col2':[    "PCR + Rs Left",
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R")'.format(start=start,end=end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R")'.format(start=start,end=end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R")'.format(start=start,end=end,location=york)],
        'col3':[    "Rapids Left",
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid")'.format(start=start,end=end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid")'.format(start=start,end=end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid")'.format(start=start,end=end,location=york)],
        'col4':[    "Total Scheduled Left",
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}")'.format(start=start,end=end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}")'.format(start=start,end=end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}")'.format(start=start,end=end,location=york) ]
    }

frame = pd.DataFrame(data=d)

print(frame)
frame.to_csv("googlesheetformulas.csv")
