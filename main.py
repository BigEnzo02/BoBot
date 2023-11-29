import os
import datetime
import asyncio
from inspect import isclass

#google api things that are probably important
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

#discord api things for bot login and use
import discord
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

#id of spreadsheet to open
#K_ID = "1ZAgeRwCdABwFE0jzJ5sgvGxBAWuCHmHiPs6tyP1jbYM"       #kirkland spreadsheet copy id
K_ID = "1mBnfaKtLnigf-_thddxbZJyynnUpeWL5GiRQa2vK8VE"       #live kirkland spreadsheet id
#W_ID = "1g0dqif8ab-Z0qe8Tyxd5rH2tqJZWjeVCb1NwMmJ1vy4"       #woodinville spreadsheet copy id
W_ID = "1mxUYNCQ1e5drvBoBFwksm7b_8fE3l8NY8m3INPdJAOM"       #live woodinville spreadsheet id

def manage_sheets(location:str):
    """Open google sheet and read values in comparison to lows

    Returns:
        lows (str): formatted message of all rows that are low
    """
    #bunch of code that basically means you don't need to sign in every time it is run
    credentials = None
    if os.path.exists("gtoken.json"):
        #token exists, read it into memory
        credentials = Credentials.from_authorized_user_file("gtoken.json", SCOPES)
    if not credentials or not credentials.valid:
        #credentials either don't exist or have expired
        if (credentials and credentials.expired and credentials.refresh_token):
            #expired credentials
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("gtoken.json", "w") as token:
            token.write(credentials.to_json())

    try:
        #try to get the spreadsheet and read values
        
        #----IMPORTANT BASICS----
        # to get specific variable value, it is easiest to use f string like f"TestSheet!A{row}"
        # also, if something is not showing up in sheets remember to use .execute()

        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()
        
        empty = 0       #count how many blank spaces consecutively
        row = 7         #starting row for count
        day = (datetime.date.today().today().weekday() + 1) % 7     #convert day into number and format it from sunday-saturday instead of monday-sunday
        letter = chr(99 + day).upper()      #convert number back into a letter associated with the correct row using ASCII, starting from row C
        acc = "Lows counts for %s %s:\n" % (location, datetime.date.today())        #accumulator string
        add_spacing = False         #line breaks should be added after empty spaces
        #is it better for this to send the day it is checking or the date written on the spreadsheet? idk

        if location[0] == 'K':
            SPREADSHEET_ID = K_ID
        elif location[0] == 'W':
            SPREADSHEET_ID = W_ID
        else:
            return 'An internal error has occured'

        while empty < 3:        #pull the next 100 rows and sort them until an empty section has been found

            #it is important to request these all at once and sort through them later to avoid passing max requests, along with being much faster
            result  =  sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"B{row}:B{row+100}").execute().get("values", str)       #low count values
            values  =  sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"{letter}{row}:{letter}{row+100}").execute().get("values", str)     #actual recorded values
            names   =  sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"A{row}:A{row+100}").execute().get("values", str)       #names of values so humans understand

            #TODO if numbers are entered in place of an "N" or "Y", it does not count as being low or throw an error
            for low in result:
                if empty < 3:        #needs to be at least 2 for kirkland and at least 3 but under 4 for woodinville
                    #low_value = result.get("values", str)
                    
                    if len(low) > 0:
                        #found full box, reset empty counter and parse
                        empty = 0

                        low_num = ""
                        for char in low[0]:        #due to the way the api is designed, each value comes within a list and must be extracted
                            if char.isdigit() or char == ".":
                                #extract number from the low count
                                low_num += char

                        try:
                            actual_num = values[row-7][0]       #subtract 7 because it is the starting row and list index starts from 0
                        except IndexError:
                            #if something is added to the sheet but does not have a value associated with it
                            acc += "Row %s is incomplete (has no value)!\n" % row
                            row += 1
                            continue
                        name = names[row-7][0]
                        try:
                            low_num = float(low_num)
                        except ValueError:
                            #low value is entered incorrectly
                            # I believe the only case this can happen is if there are multiple decimal points due to earlier check
                            acc += "Invalid num for %s: have %s (need  %s [invalid])\n" % (name, actual_num, low_num)
                            row += 1
                            continue

                        if actual_num.upper() == "Y":
                            pass
                        elif actual_num.upper() == "N":
                            #low on a non-numerical value like cups or lids
                            acc += ("%s: have none (need %s)\n" % (name, low_num))
                        else:

                            try:
                                actual_num = float(actual_num)
                            except ValueError:
                                #actual number is entered incorrectly
                                acc += "Invalid num for %s: have %s [invalid] (need  %s)\n" % (name, actual_num, low_num)
                                row += 1
                                continue

                            if actual_num < low_num:
                                if add_spacing:     #this adds a space if there is an empty section but avoids adding three extra spaces at the end
                                    #TODO does not add spaces in case of error (should it?)
                                    acc += '\n'
                                    add_spacing = False

                                acc += ("%s: have %s (need %s)\n" % (name, actual_num, low_num))

                    else:
                        #an empty box returns <class 'str'>
                        empty += 1
                        add_spacing = True      #add a newline character on the next go-around
                    
                    row += 1
                else:
                    #found 4 consecutive empty spaces, return the result
                    return acc
        return acc      #in case there are less than 4 empty spaces at the end of the count (values 7-104 are used) 
    except HttpError as e:
        print("HTTP error while reading spreadsheet: %s" % e)
        return "An internal error has occured: see below for details:\n%s" % e

client = discord.Client()

@client.event
async def on_ready():
    print(f'{client.user} has connected to Discord!')

@client.event
async def on_message(message):
    if message.author == client.user:
        #if it is a message sent by the bot, return
        return

    if message.content.startswith('$'):        #public commands

        channel = client.get_channel(message.channel.id)
            
        if message.content.startswith('$ping'):
            await channel.send('pong!') #type: ignore
        if message.content.startswith('$low'):
            
            if message.channel.id == 889600873880223785:        #woodinville low counts channel
                await channel.send(manage_sheets(location='Woodinville')) #type: ignore
            elif message.channel.id == 935324349169299536:      #kirkland low counts channel
                await channel.send(manage_sheets(location='Kirkland')) #type: ignore
            else:
                
                #either someone sent a message to the wrong channel or channel ID has changed
                await channel.send('Which location would you like to count?') #type: ignore
                msg = ''
                author = message.author     #so it waits for only the author of the message and does not get messed with by other people
                try:
                    msg = await client.wait_for('message', check=lambda message: message.author == author, timeout=20)
                except asyncio.TimeoutError:
                    pass        #no message sent
                if msg.content.lower() == 'woodinville': #type: ignore
                    await channel.send(manage_sheets(location='Woodinville')) #type: ignore
                elif msg.content.lower() == 'kirkland': #type: ignore
                    print("Pulling '%s' results for user %s" % (message.content, message.author))
                    await channel.send(manage_sheets(location='Kirkland')) #type: ignore
                else:
                    await channel.send("Sorry, I don't know that location!") #type: ignore


async def get_output_task():

    await client.wait_until_ready()
    channel = client.get_channel(1177305755985125447)#testing channel
    await channel.send('up and running!') #type: ignore         #send a message to internal testing channel

#print(manage_sheets())
client.loop.create_task(get_output_task())
client.run('MTE3NzMwMTc0ODcxNzcxNTQ1Ng.GWN-3p.zqq9LYaR7VwBJBlS676PB4QdqGSi_i5jlM9znM')