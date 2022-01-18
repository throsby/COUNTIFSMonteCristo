import pandas as pd
import os
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from gspread_pandas import Spread, Client
import gspread_formatting
from oauth2client.service_account import ServiceAccountCredentials
from sheets_pdfreader import open_in_chrome

start = "4"
end = "178"
top_left = 180
add_start = str(top_left+27)
add_end = str(int(add_start)+20)
york = "York*"
wash = "Washington*"
set = "Set*"
on_set=False
locations = [york,wash,set]

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/Throsby/Documents/GitHub/PDF_Reader/secret_key.json', scope)
client = gspread.authorize(creds)


setwashyork  = { 'colLoc':[      "Scheduled","Set","Washington","York","","","Not Yet Performed","Set","Washington","York","","","Completed","Set","Washington","York","","","Cancelled","Set","Washington","York",""],
                 'colPCR':[      "PCRs",
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="PCR"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR"),
                                 "=SUM($E${top}:$E${bot})".format(top=top_left+1,bot=top_left+3),
                                 "",
                                 "PCRs Left",
                                 '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="PCR"),
                                 '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR"),
                                 '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR"),
                                 "=SUM($E${top}:$E${bot})".format(top=top_left+7,bot=top_left+9),
                                 "",
                                 "PCRs Completed",
                                 '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="PCR"),
                                 '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR"),
                                 '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR"),
                                 "=SUM($E${top}:$E${bot})".format(top=top_left+13,bot=top_left+15),
                                 "",
                                 "PCRs Cancelled",
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="PCR"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR"),
                                 "=SUM($E${top}:$E${bot})".format(top=top_left+19,bot=top_left+21)],
                  'colRPCR':[    "PCR + Rs",
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="PCR + R"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR + R"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR + R"),
                                 "=SUM($F${top}:$F${bot})".format(top=top_left+1,bot=top_left+3),
                                 "",
                                 "PCR + Rs Left",
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"{type}") - COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="PCR + R"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"{type}") - COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR + R"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"{type}") - COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR + R"),
                                 "=SUM($F${top}:$F${bot})".format(top=top_left+7,bot=top_left+9),
                                 "",
                                 "PCR + Rs Completed",
                                 '=COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="PCR + R"),
                                 '=COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR + R"),
                                 '=COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR + R"),
                                 "=SUM($F${top}:$F${bot})".format(top=top_left+13,bot=top_left+15),
                                 "",
                                 "PCR + Rs Cancelled",
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="PCR + R"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR + R"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR + R"),
                                 "=SUM($F${top}:$F${bot})".format(top=top_left+19,bot=top_left+21)],
                  'colRapid':[   "Rapids",
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="Rapid"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="Rapid"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="Rapid"),
                                 "=SUM($G${top}:$G${bot})".format(top=top_left+1,bot=top_left+3),
                                 "",
                                 "Rapids Left",
                                 '=COUNTIFS($B${start}:$B${end}, "<>r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($B${add_start}:$B${add_end}, "<>r", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="Rapid"),
                                 '=COUNTIFS($B${start}:$B${end}, "<>r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($B${add_start}:$B${add_end}, "<>r", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="Rapid"),
                                 '=COUNTIFS($B${start}:$B${end}, "<>r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($B${add_start}:$B${add_end}, "<>r", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="Rapid"),
                                 "=SUM($G${top}:$G${bot})".format(top=top_left+7,bot=top_left+9),
                                 "",
                                 "Rapids Completed",
                                 '=COUNTIFS($B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="Rapid"),
                                 '=COUNTIFS($B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="Rapid"),
                                 '=COUNTIFS($B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="Rapid"),
                                 "=SUM($G${top}:$G${bot})".format(top=top_left+13,bot=top_left+15),
                                 "",
                                 "Rapids Cancelled",
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=set,type="Rapid"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="Rapid"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="Rapid"),
                                 "=SUM($G${top}:$G${bot})".format(top=top_left+19,bot=top_left+21)],
                 'colTotal':[    "Total Tests Scheduled",
                                 "=SUM($E${top}:$G${top})".format(top=top_left+1),
                                 "=SUM($E${mid}:$G${mid})".format(mid=top_left+2),
                                 "=SUM($E${bot}:$G${bot})".format(bot=top_left+3),
                                 "=SUM($E${tot}:$G${tot})".format(tot=top_left+4),
                                 "",
                                 "Total Yet to Be Performed",
                                 "=SUM($E${top}:$G${top})".format(top=top_left+7),
                                 "=SUM($E${mid}:$G${mid})".format(mid=top_left+8),
                                 "=SUM($E${bot}:$G${bot})".format(bot=top_left+9),
                                 "=SUM($E${tot}:$G${tot})".format(tot=top_left+10),
                                 "",
                                 "Total Tests Completed",
                                 "=SUM($E${top}:$G${top})".format(top=top_left+13),
                                 "=SUM($E${mid}:$G${mid})".format(mid=top_left+14),
                                 "=SUM($E${bot}:$G${bot})".format(bot=top_left+15),
                                 "=SUM($E${tot}:$G${tot})".format(tot=top_left+16),
                                 "",
                                 "Total Tests Cancelled",
                                 "=SUM($E${top}:$G${top})".format(top=top_left+19),
                                 "=SUM($E${mid}:$G${mid})".format(mid=top_left+20),
                                 "=SUM($E${bot}:$G${bot})".format(bot=top_left+21),
                                 "=SUM($E${tot}:$G${tot})".format(tot=top_left+22)]
     }

washyork     = { 'colLoc':[      "Scheduled","Washington","York","","","Not Yet Performed","Washington","York","","","Completed","Washington","York","","","Cancelled","Washington","York",""],
                 'colPCR':[      "PCRs",
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR"),
                                 "=SUM($E${top}:$E${bot})".format(top=top_left+1,bot=top_left+2),
                                 "",
                                 "PCRs Left",
                                 '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR"),
                                 '=COUNTIFS($C${start}:$C${end}, "<>y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR"),
                                 "=SUM($E${top}:$E${bot})".format(top=top_left+6,bot=top_left+7),
                                 "",
                                 "PCRs Completed",
                                 '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR"),
                                 '=COUNTIFS($C${start}:$C${end}, "y", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR"),
                                 "=SUM($E${top}:$E${bot})".format(top=top_left+11,bot=top_left+12),
                                 "",
                                 "PCRs Cancelled",
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR"),
                                 "=SUM($E${top}:$E${bot})".format(top=top_left+16,bot=top_left+17)],
                  'colRPCR':[    "PCR + Rs",
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR + R"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR + R"),
                                 "=SUM($F${top}:$F${bot})".format(top=top_left+1,bot=top_left+2),
                                 "",
                                 "PCR + Rs Left",
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end},"{type}") - COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR + R"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR + R"),
                                 "=SUM($F${top}:$F${bot})".format(top=top_left+6,bot=top_left+7),
                                 "",
                                 "PCR + Rs Completed",
                                 '=COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR + R"),
                                 '=COUNTIFS($C${start}:$C${end}, "y", $B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($C${add_start}:$C${add_end}, "y", $B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR + R"),
                                 "=SUM($F${top}:$F${bot})".format(top=top_left+11,bot=top_left+12),
                                 "",
                                 "PCR + Rs Cancelled",
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="PCR + R"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="PCR + R"),
                                 "=SUM($F${top}:$F${bot})".format(top=top_left+16,bot=top_left+17)],
                  'colRapid':[   "Rapids",
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="Rapid"),
                                 '=COUNTIFS($K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="Rapid"),
                                 "=SUM($G${top}:$G${bot})".format(top=top_left+1,bot=top_left+2),
                                 "",
                                 "Rapids Left",
                                 '=COUNTIFS($B${start}:$B${end}, "<>r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="Rapid"),
                                 '=COUNTIFS($B${start}:$B${end}, "<>r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") - COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="Rapid"),
                                 "=SUM($G${top}:$G${bot})".format(top=top_left+6,bot=top_left+7),
                                 "",
                                 "Rapids Completed",
                                 '=COUNTIFS($B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="Rapid"),
                                 '=COUNTIFS($B${start}:$B${end}, "r", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($B${add_start}:$B${add_end}, "r", $K${add_start}:$K${add_end}, "{location}",$L${add_start}:$L${add_end},"{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="Rapid"),
                                 "=SUM($G${top}:$G${bot})".format(top=top_left+11,bot=top_left+12),
                                 "",
                                 "Rapids Cancelled",
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=wash,type="Rapid"),
                                 '=COUNTIFS($A${start}:$A${end}, "canc*", $K${start}:$K${end}, "{location}", $L${start}:$L${end},"{type}") + COUNTIFS($A${add_start}:$A${add_end}, "canc*", $K${add_start}:$K${add_end}, "{location}", $L${add_start}:$L${add_end}, "{type}")'.format(start=start,end=end,add_start=add_start,add_end=add_end,location=york,type="Rapid"),
                                 "=SUM($G${top}:$G${bot})".format(top=top_left+16,bot=top_left+17)],
                 'colTotal':[    "Total Tests Scheduled",
                                 "=SUM($E${top}:$G${top})".format(top=top_left+1),
                                 "=SUM($E${mid}:$G${mid})".format(mid=top_left+2),
                                 "=SUM($E${bot}:$G${bot})".format(bot=top_left+3),
                                 "",
                                 "Total Left",
                                 "=SUM($E${top}:$G${top})".format(top=top_left+6),
                                 "=SUM($E${mid}:$G${mid})".format(mid=top_left+7),
                                 "=SUM($E${bot}:$G${bot})".format(bot=top_left+8),
                                 "",
                                 "Total Tests Completed",
                                 "=SUM($E${top}:$G${top})".format(top=top_left+11),
                                 "=SUM($E${mid}:$G${mid})".format(mid=top_left+12),
                                 "=SUM($E${bot}:$G${bot})".format(bot=top_left+13),
                                 "",
                                 "Total Tests Cancelled",
                                 "=SUM($E${top}:$G${top})".format(top=top_left+16),
                                 "=SUM($E${mid}:$G${mid})".format(mid=top_left+17),
                                 "=SUM($E${bot}:$G${bot})".format(bot=top_left+18)]
     }


try:
    sheet = client.open("Tables for the Sheet")
except:
    sheet = client.create("Tables for the Sheet")
worksheet = sheet.sheet1

if(on_set==True):
    frame = pd.DataFrame(data=setwashyork)
    #Side Lines
    worksheet.format("A3:A5",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("A9:A11",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("A15:A17",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("A21:A23",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("E3:E5",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("E9:E11",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("E15:E17",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("E21:E23",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    #Top and Bottom Lines
    worksheet.format("B2:D2",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B6:D6",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B8:D8",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B12:D12",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B14:D14",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B18:D18",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B20:D20",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B24:D24",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    #Corners
    worksheet.format("A2",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E2",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A6",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E6",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A8",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E8",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A12",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E12",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A14",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E14",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A18",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E18",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A20",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E20",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A24",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E24",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})

    ###---------------------------Send to CSV File-------------------------------###
    print("Add start: {st} Add end: {end}".format(st=add_start,end=add_end))
    frame.to_csv("./outgooglesheetformulas.csv")

    ###------------------------Send to Google Sheets-----------------------------###
    set_with_dataframe(worksheet,frame)

    worksheet.format("B3:E6",{"horizontalAlignment":"CENTER"})
    worksheet.format("B9:E12",{"horizontalAlignment":"CENTER"})
    worksheet.format("B15:E18",{"horizontalAlignment":"CENTER"})
    worksheet.format("B21:E24",{"horizontalAlignment":"CENTER"})

elif(on_set==False):
    frame = pd.DataFrame(data=washyork)
    #Format grids for 4x4 tables
    #Side Lines
    worksheet.format("A3:A4",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("A8:A9",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("A13:A14",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("A18:A19",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("E3:E4",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("E8:E9",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("E13:E14",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("E18:E19",{"borders":{"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    #Top and Bottom Lines
    worksheet.format("B2:D2",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B5:D5",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B7:D7",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B10:D10",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B12:D12",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B15:D15",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B17:D17",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    worksheet.format("B20:D20",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0},}}})
    #Corners
    worksheet.format("A2",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E2",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A5",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E5",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A7",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E7",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A10",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E10",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A12",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E12",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A15",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E15",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A17",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E17",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("A20",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})
    worksheet.format("E20",{"borders":{"top":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"bottom":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"left":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}},"right":{"style":"SOLID","color":{"red":0,"green":0,"blue":0}}}})

    ###---------------------------Send to CSV File-------------------------------###
    print("Add start: {st} Add end: {end}".format(st=add_start,end=add_end))
    frame.to_csv("./outgooglesheetformulas.csv")

    ###------------------------Send to Google Sheets-----------------------------###
    set_with_dataframe(worksheet,frame)

    worksheet.format("B2:E6",{"horizontalAlignment":"CENTER"})
    worksheet.format("B9:E12",{"horizontalAlignment":"CENTER"})
    worksheet.format("B15:E18",{"horizontalAlignment":"CENTER"})
    worksheet.format("B21:E24",{"horizontalAlignment":"CENTER"})
open_in_chrome()
