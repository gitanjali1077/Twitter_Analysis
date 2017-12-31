import tweepy
import csv
import argparse
import parse
from datetime import datetime, timedelta
import datetime
import xlsxwriter
from subprocess import Popen


consumer_key = 'msm3MlyIy2K86YWbCoNr5Ccg4'
consumer_secret = '0R1HCdmzpyQba3wjrRQmwtuW52zutnFQslRbSpy6TxM3eC1fJM'
access_token = '4641961574-VRRgCFhLti71Xn1POww21T3czPMhANCYAMN4lTq'
access_secret = '9jnaMqFjWoW5A9KMY1DB2c1bvgrgHBtzITon9d5G9kvAE'

def tw_parser():
    global qw, ge, l, t, c, d,qw1

# USE EXAMPLES:
# =-=-=-=-=-=-=
# % twsearch <search term>            --- searches term
# % twsearch <search term> -g sf      --- searches term in SF geographic box <DEFAULT = none>
# % twsearch <search term> -l en      --- searches term with lang=en (English) <DEFAULT = en>
# % twsearch <search term> -t {m,r,p} --- searches term of type: mixed, recent, or popular <DEFAULT = recent>
# % twsearch <search term> -c 12      --- searches term and returns 12 tweets (count=12) <DEFAULT = 1>
# % twsearch <search term> -o {ca, tx, id, co, rtc)   --- searches term and sets output options <DEFAULT = ca, tx>

# Parse the command
    parser = argparse.ArgumentParser(description='Twitter Search')
    parser.add_argument(action='store', dest='query', help='Search term string')
    parser.add_argument(action='store', dest='query1', help='Search term string second file')
    args = parser.parse_args()

    qw = args.query     # Actual query word(s)
    qw1= args.query1
    ge = ''
    #print ("Query: %s, Location: %s, Language: %s, Search type: %s, Count: %s" %(qw,ge,l,t,c))


def main():
  tw_parser()
  auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
  auth.set_access_token(access_token, access_secret)
  api = tweepy.API(auth)

  final=0

  workbook = xlsxwriter.Workbook('Result.xlsx')
  worksheet = workbook.add_worksheet()
  bold = workbook.add_format({'bold': 1})
  headings = ['Hours', qw, qw1]
  array1=[]
  array2=[]
# Add the worksheet data that the charts will refer to.

# i is here maintaining the hours, It is starting from 6 bcz we cant retrieve data of last
# 5.5 hours from tweepy api
  #for i in range(6,8):
  each_count=0
  i=6
  for tweet in tweepy.Cursor(api.search,q=qw,count=100,\
                           lang="en").items():
     a=  datetime.datetime.now() - timedelta(hours=i) #days=1)
     if   tweet.created_at > a:
      each_count=each_count+1
      final=final+1
      #csvWriter.writerow([ tweet.created_at, tweet.text.encode('utf-8'),tweet.retweet_count,final-each_count,final])
      #print tweet.created_at , tweet.text.encode('utf-8'), each_count,final# ,tweet.retweet_count
      #print tweet.follower_count
     else:
         if i==15:
             array1.append(final)
             each_count=0
             break
         else:
             array1.append(final)
             each_count=0
             i=i+1
  # csvWriter.writerow([tweet.created_at, final-each_count,final])
  #array1.append(final)
   
  

  print "for first personality "+ qw


  final1=0
  i=6
  #for i in range(6,8):
  each_count1=0   
  for tweet in tweepy.Cursor(api.search,q=qw1,count=100,\
                           lang="en").items():
     a=  datetime.datetime.now() - timedelta(hours=i) #days=1)
     if   tweet.created_at > a:
      each_count1=each_count1+1
      final1=final1+1
      #print tweet.created_at , tweet.text.encode('utf-8'), each_count1,final1# ,tweet.retweet_count
     else:
        if i==15:
             array2.append(final1)
             each_count1=0
             break
        else:
             array2.append(final1)
             each_count1=0
             i=i+1 
   #csvWriter1.writerow([tweet.created_at,i, final1-each_count1,final1])
   #array2.append(final1)  
  

  print "for second personality " + qw1
  
 # print (row_count)
# excel pe working
  data = [
    [1,2, 3, 4, 5, 6, 7,8,9,10,11],
   # [10, 40, 50, 20, 10, 50],
    array1,
    array2,
  ]
  worksheet.write_row('A1', headings, bold)
  worksheet.write_column('A2', data[0])
  worksheet.write_column('B2', data[1])
  worksheet.write_column('C2', data[2])

#######################################################################
#
# Create a new bar chart.
#
  chart1 = workbook.add_chart({'type': 'bar'})

# Configure the first series.
  chart1.add_series({
    'name':       '=Sheet1!$B$1',
    'categories': '=Sheet1!$A$2:$A$12',
    'values':     '=Sheet1!$B$2:$B$12',
  })

# Configure a second series. Note use of alternative syntax to define ranges.
  chart1.add_series({
    'name':       ['Sheet1', 0, 2],
    'categories': ['Sheet1', 1, 0, 10, 0],
    'values':     ['Sheet1', 1, 2, 10, 2],
  })

# Add a chart title and some axis labels.
  chart1.set_title ({'name': 'Results of analysis'})
  chart1.set_x_axis({'name': 'Tweets count'})
  chart1.set_y_axis({'name': 'Number of hours'})

# Set an Excel chart style.
  chart1.set_style(11)

# Insert the chart into the worksheet (with an offset).
  worksheet.insert_chart('D2', chart1, {'x_offset': 25, 'y_offset': 10})


  workbook.close()
  p = Popen('Result.xlsx', shell=True)

  

if __name__ == "__main__":
    main()
