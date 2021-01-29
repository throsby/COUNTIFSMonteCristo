import pandas as pd
import os

start = "4"
end = "306"
add_start = "335"
add_end = "355"
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

# Including additionals column
d2 = {  'col0':["","Set","Washington","York"],
        'col1':[    "PCRs Left Done",
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york)],
        'col2':[    "PCR + Rs Done",
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR + R")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR + R")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR + R")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york)],
        'col3':[    "Rapids Done",
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"Rapid")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"Rapid")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"Rapid")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york)],
        'col4':[    "Total Tests Done",
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york)]
    }

d3 = {  'col0':["","","Set","Washington","York"],
        'col1':[    "PCRs Left Done",
                    '',
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york)],
        'col2':[    "PCR + Rs Done",
                    '',
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR + R")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR + R")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"PCR + R") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"PCR + R")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york)],
        'col3':[    "Rapids Done",
                    '',
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"Rapid")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"Rapid")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"Rapid") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"Rapid")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york)],
        'col4':[    "Total Tests Done",
                    '',
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash),
                    '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york)]
    }

frame = pd.DataFrame(data=d2)

print(frame)
frame.to_csv("./googlesheetformulas.csv")
