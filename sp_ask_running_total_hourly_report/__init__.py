__version__ = '0.1.5'

#std library
from collections import Counter
from collections import OrderedDict
import calendar
import datetime
from pprint import pprint as print

#packages
import pandas as pd
import arrow
import lh3.api as lh3


client = lh3.Client()
chats = client.chats()

def create_excel_file(filename, data):
    df = pd.DataFrame(data)

    #reorder columns
    cols = df.columns.tolist()
    del cols[cols.index('date')]
    cols = ['date'] + cols
    df = df[cols]
    #df = df.rename({'date': 'day of the month'}, axis=1)

    writer = pd.ExcelWriter(filename+".xlsx", engine='xlsxwriter')
    df.to_excel(writer, index=False)  # send df to writer
    worksheet = writer.sheets['Sheet1']  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
    writer.save()
    print(filename+".xlsx")

def get_answered_chats(chats):
    chat_not_none = [chat for chat in chats if chat.get("accepted") != None]
    return chat_not_none

def get_unanswered_chats(chats):
    chat_unanswered = [chat for chat in chats if chat.get("accepted") == None]
    return chat_unanswered

def remove_practice_queues(chats_this_day):
    res = [chat for chat in chats_this_day if not "practice" in chat.get("queue")]
    return res

def create_report(year=2019, month=2):
    given_date = datetime.datetime(year, month, 1)
    month = given_date.month
    month_name = given_date.strftime("%B")
    filename = str(month) + "-" + month_name

    first_day, last_day = calendar.monthrange(year, month)

    total_answered_chats = []
    days = []
    data = []
    chats_per_hour = list()
    for day in range(1,last_day+1):
        all_chats = chats.list_day(year,month,day)
        chats_this_day = remove_practice_queues(all_chats)
        chat_not_none = [chat for chat in chats_this_day if chat.get("accepted") != None]
    
        list_of_hours = list()
        for chat in chat_not_none:
            this_date = arrow.get(chat.get('accepted'))
            list_of_hours.append(this_date.hour)
        
        result = dict(Counter(list_of_hours))
        result = dict(sorted(result.items()))
        this_day = str(year) + '-'+ str(month) + '-'+ str(day)
        print(this_day)
        result['date'] = this_day
        chats_per_hour.append(result)
    res = pd.DataFrame(chats_per_hour)
    #print(res)
    
    create_excel_file(filename, chats_per_hour)


if __name__ == '__main__':
    #create_report(2019, 11)

    for month_number in range(1, 13):
        #pass
        create_report(2019, month_number)

