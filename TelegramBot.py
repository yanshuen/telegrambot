from operator import contains
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, Update, Bot
from telegram.ext import (
    Application,
    CommandHandler,
    ConversationHandler,
    MessageHandler,
    filters
)
import time
import logging
from datetime import date, datetime, timedelta
import os
import asyncio
import requests
import pymongo
import pyodbc
import uuid
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import json


##mongodb db
#myclient = pymongo.MongoClient("mongodb://localhost:27017/")
#mydb = myclient["telegram_bot_chatids"]
#mycol = mydb["chatids"]


#sqlserver db
driver = "ODBC Driver 18 for SQL Server"
server = "DESKTOP-S0MOAN3\SQLEXPRESS02"
database = "telegrambot"
userid = "sa"
password = "Password1"
connection = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={userid};PWD={password};TrustServerCertificate=yes')
cursor = connection.cursor()


#json conf file
json_conf_file_path = os.path.abspath("bot_conf.json")
file = open(json_conf_file_path, "r")
data = json.load(file)
keyboards = data["keyboard"]


current_date = datetime.now().strftime("%Y%m%d")

#py log
logging.basicConfig(filename=data["python_log"].format(conf_py_current_date=current_date), level=logging.ERROR)

#tele bot log
file_dir = data["bot_log"].format(conf_bot_current_date=current_date)


#ys token: 1147547785
#ck token: 461959755
#gc token: -4064577558

#query mongodb for ids for broadcast msg
#broadcast_chat_id = []
#all_documents = mycol.find()
#for document in all_documents:
#    for key in document:
#        if key != '_id':
#            broadcast_chat_id.append(key)


#broadcast msg
broadcast_chat_id = []
cursor.execute('SELECT [chat_id] FROM [telegrambot].[dbo].[bot_details]')
broadcast_msg_ids = cursor.fetchall()
for row in broadcast_msg_ids:
    for x in row:
        if x not in broadcast_chat_id:
            broadcast_chat_id.append(x)

token = "6364510958:AAHghv-nCZ8R2eZcivuiFNNFZU-IQZ2Boas"
broadcast_msg = "This is a broadcast message."

#if len(broadcast_chat_id) != 0:
#    for chat_id in broadcast_chat_id:
#        to_url = 'https://api.telegram.org/bot{}/sendMessage?chat_id={}&text={}&parse_mode=HTML'.format(token, chat_id, broadcast_msg)
#        requests.get(to_url)


user_chat_id = 0
user_question = ""

editName, editUnit, startUnit, startName, issueChoice, tabletIssueChoice, appIssueChoice, systemIssueChoice, userenquiryIssueChoice, newQuestion, solveProblem = range(11)

async def start(update, context):
    #for bot details log 
    global chat_name
    global current_datetime

    #for sqlserver
    global sql_chat_id
    global sql_date
    global sql_time
    global sql_unit
    global sql_reported_by
    global sql_category
    global sql_issue_reported
    global sql_response_time
    global sql_date_closed
    global sql_status
    global sql_remarks 

    global db_unit
    global db_reported_by

    #sqlserver variables 
    sql_chat_id = None
    sql_date = None
    sql_time = None
    sql_unit = None
    sql_reported_by = None
    sql_category = None
    sql_issue_reported = None
    sql_response_time = None
    sql_date_closed = None
    sql_status = None
    sql_remarks = None

    chat = update.message.chat
    if chat.username == None:
        if chat.first_name == None:
            chat_name = repr(chat.last_name).strip("'")
        elif chat.last_name == None:
            chat_name = repr(chat.first_name).strip("'")
        else:
            chat_name = repr(chat.first_name).strip("'") + " " + repr(chat.last_name).strip("'")
    else:
        chat_name = repr(chat.username).strip("'")

    user = update.message.from_user

    message = update.message
    
    #bot details log
    if not os.path.isdir(file_dir):
        os.mkdir(file_dir) 
    current_datetime = datetime.now().strftime('%d %B %Y_%H.%M.%S')
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a", encoding='utf-8')

    log_current_date = datetime.now().strftime('%d/%m/%Y')
    log_file.write(log_current_date + "\n")
    log_file.write("User/Chat ID: " + repr(chat.id) + "\n\n\n")

    log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")


    #setting value for sql db
    sql_chat_id = chat.id
    sql_date = message.date + timedelta(hours=8)
    sql_time = message.date + timedelta(hours=8)


    #checking if id exists in sql server
    cursor.execute(f'select unit, reported_by from [dbo].[bot_details] where chat_id = {sql_chat_id}')
    check_id_exists_result = cursor.fetchone()

    #if return none means the id doesnt exist, if it returns unit means that the id is alr in db n it cld retrieve the unit value 
    if check_id_exists_result != None:
        #exists in db

        db_unit = check_id_exists_result[0]
        db_reported_by = check_id_exists_result[1]

        check_nameunit_keyboard = [["Yes"], ["No"]]

        #keyboard formatting
        formatted_list = [f"['{item}']" for inner_list in check_nameunit_keyboard for item in inner_list]
        new_check_nameunit_keyboard = ', '.join(formatted_list)

        reply_markup = ReplyKeyboardMarkup(check_nameunit_keyboard, resize_keyboard=True, one_time_keyboard=True)

        reply = await update.message.reply_text("Welcome to S3 Bot! 🤖 I am your personal assistant for today. \n\n"
                                        "Before we start, please help me to confirm some of your personal details. 📃 \n\n"
                                        "Is the following correct? \n"
                                        "Full name: " + check_id_exists_result[1] + " \n"
                                        "Unit: " + check_id_exists_result[0], 
                                        reply_markup=reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_check_nameunit_keyboard).strip('"') + "\n\n")
    

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))

        return editName

    else:
        #doesnt exist in db

        reply = await update.message.reply_text("Welcome to S3 Bot! 🤖 I am your personal assistant for today. \n\n"
                                                "Before we start, I would need some of your personal details. 📃\n\n"
                                                "Please help me key in your full name.", 
                                                reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))

        return startUnit

    ##checking if id exists in mongodb 
    #chat_id_string = str(chat.id)
    #check_query = {chat_id_string: {"$exists": True}}    
    #count_result = mycol.count_documents(check_query)
    #if count_result > 0:
    #    #id exists in db

        #start_keyboard = [["Tablet Issues"], ["Application System Issues"], ["Others"]]

        ##keyboard formatting
        #formatted_list = [f"['{item}']" for inner_list in start_keyboard for item in inner_list]
        #new_start_keyboard = ', '.join(formatted_list)

        #reply_markup = ReplyKeyboardMarkup(start_keyboard, resize_keyboard=True, one_time_keyboard=True)

        #reply = await update.message.reply_text("Welcome to S3 Bot! I am your personal assistant for today. \n\n"
        #                                        "Click any of the options below as to which problem you are facing, and I will do our best to help! \n\n\n"
        #                                        "If at any time you would like to end the conversation /cancel, or start a new conversation /start, head over to the menu tab to view the list of commands! ", 
        #                                        reply_markup=reply_markup)


        ##bot details log
        #bot_current_date = datetime.now().strftime("%H:%M:%S")
        #log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        #log_file.write("Keyboard Options: " + repr(new_start_keyboard).strip('"') + "\n\n")
    

        #log_file.close()

        ##create timer using asyncio create task
        #context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))

        #return issueChoice


    #else:
    #    #doesnt exist in db

        #reply = await update.message.reply_text("Welcome to S3 Bot! I am your personal assistant for today. \n\n"
        #                                        "Before we start, I would need some of your personal details. \n\n"
        #                                        "Please help me key in your full name.", 
        #                                        reply_markup=ReplyKeyboardRemove())


        ##bot details log
        #bot_current_date = datetime.now().strftime("%H:%M:%S")
        #log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")

        #log_file.close()

        ##create timer using asyncio create task
        #context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))

        #return startUnit

async def edit_name(update, context):
    global current_datetime
    global sql_unit
    global sql_reported_by

    #global db_unit
    #global db_reported_by

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")
    
    if (update.message.text) == "No":
        #need to edit info 

        #restart timer
        context.user_data['timer_task'].cancel()

        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Please key in your full name.", 
                                    reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))

        return editUnit
        
    else:
        #dont need to edit info

        #restart timer
        context.user_data['timer_task'].cancel()

        #setting value for sql db 
        sql_unit = db_unit
        sql_reported_by = db_reported_by

        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")
        
        start_keyboard = [["Hardware"], ["Software"], ["System"], ["User Enquiry"]]

        #keyboard formatting
        formatted_list = [f"['{item}']" for inner_list in start_keyboard for item in inner_list]
        new_start_keyboard = ', '.join(formatted_list)

        reply_markup = ReplyKeyboardMarkup(start_keyboard, resize_keyboard=True, one_time_keyboard=True)

        reply = await update.message.reply_text("Click any of the options below as to which category of problem you are facing.\n\n\n"
                                        "If you would like to end the conversation /cancel, or start a new conversation /start, head over to the menu tab to view the list of commands. ", 
                                        reply_markup=reply_markup)


        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_start_keyboard).strip('"') + "\n\n")
        
        log_file.close()

        return issueChoice


async def edit_unit(update, context):
    #restart timer
    context.user_data['timer_task'].cancel()

    global current_datetime
    global sql_reported_by

    #setting value for sql db 
    sql_reported_by = update.message.text

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

    reply = await update.message.reply_text("Please key in your unit.", 
                                    reply_markup=ReplyKeyboardRemove())

    #bot details log
    bot_current_date = datetime.now().strftime("%H:%M:%S")
    log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
    

    log_file.close()

    return startName


async def start_unit(update, context):
    #restart timer
    context.user_data['timer_task'].cancel()

    global current_datetime
    global sql_reported_by

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

    #setting value for sql db 
    sql_reported_by = message.text

    ##insert into mongodb
    #chat = update.message.chat
    #chat_id_string = str(chat.id)
    #chat_id_dict = {chat_id_string: message.text}
    #mycol.insert_one(chat_id_dict)

    reply = await update.message.reply_text("Please help me key in your unit.", 
                                    reply_markup=ReplyKeyboardRemove())


    #bot details log
    bot_current_date = datetime.now().strftime("%H:%M:%S")
    log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
    

    log_file.close()

    #create timer using asyncio create task
    context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))

    return startName
    

async def start_name(update, context):
    #restart timer
    context.user_data['timer_task'].cancel()

    global current_datetime
    global sql_unit

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

    #setting value for sql db
    sql_unit = message.text

    start_keyboard = [["Hardware"], ["Software"], ["System"], ["User Enquiry"]]

    #keyboard formatting
    formatted_list = [f"['{item}']" for inner_list in start_keyboard for item in inner_list]
    new_start_keyboard = ', '.join(formatted_list)

    reply_markup = ReplyKeyboardMarkup(start_keyboard, resize_keyboard=True, one_time_keyboard=True)

    reply = await update.message.reply_text("Thank you! \n\n"
                                    "Click any of the options below as to which category of problem you are facing, \n\n\n"
                                    "If you would like to end the conversation /cancel, or start a new conversation /start, head over to the menu tab to view the list of commands. ", 
                                    reply_markup=reply_markup)


    #bot details log
    bot_current_date = datetime.now().strftime("%H:%M:%S")
    log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
    log_file.write("Keyboard Options: " + repr(new_start_keyboard).strip('"') + "\n\n")
    

    log_file.close()

    #create timer using asyncio create task
    context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))

    return issueChoice


async def issue_choice(update, context):
    #bot details log
    global current_datetime
    global sql_category

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")


    tablet_issues_keyboard = []
    for keyboard in keyboards:
        if "hardware" in keyboard:
            tablet_issues_keyboard.append([keyboard["hardware"]])
    tablet_issues_reply_markup = ReplyKeyboardMarkup(tablet_issues_keyboard, resize_keyboard=True, one_time_keyboard=True)

    #keyboard formatting
    formatted_list = [f"['{item}']" for inner_list in tablet_issues_keyboard for item in inner_list]
    new_tablet_issues_keyboard = ', '.join(formatted_list)


    app_issues_keyboard = []
    for keyboard in keyboards:
        if "software" in keyboard:
            app_issues_keyboard.append([keyboard["software"]])
    app_issues_reply_markup = ReplyKeyboardMarkup(app_issues_keyboard, resize_keyboard=True, one_time_keyboard=True)

    #keyboard formatting
    formatted_list = [f"['{item}']" for inner_list in app_issues_keyboard for item in inner_list]
    new_app_issues_keyboard = ', '.join(formatted_list)


    system_issues_keyboard = []
    for keyboard in keyboards:
        if "system" in keyboard:
            system_issues_keyboard.append([keyboard["system"]])
    system_issues_reply_markup = ReplyKeyboardMarkup(system_issues_keyboard, resize_keyboard=True, one_time_keyboard=True)

    #keyboard formatting
    formatted_list = [f"['{item}']" for inner_list in system_issues_keyboard for item in inner_list]
    new_system_issues_keyboard = ', '.join(formatted_list)


    userenquiry_issues_keyboard = []
    for keyboard in keyboards:
        if "user_enquiry" in keyboard:
            userenquiry_issues_keyboard.append([keyboard["user_enquiry"]])
    userenquiry_issues_reply_markup = ReplyKeyboardMarkup(userenquiry_issues_keyboard, resize_keyboard=True, one_time_keyboard=True)

    #keyboard formatting
    formatted_list = [f"['{item}']" for inner_list in userenquiry_issues_keyboard for item in inner_list]
    new_userenquiry_issues_keyboard = ', '.join(formatted_list)


    if (update.message.text == "Hardware"):
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Okay, I see that you are facing " + update.message.text +
                                        " issues. Click on the option which best applies to your situation.",
                                        reply_markup=tablet_issues_reply_markup)
        
        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_tablet_issues_keyboard).strip('"') + "\n\n")

        log_file.close()

        #setting value for sql db
        sql_category = message.text

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return tabletIssueChoice


    elif (update.message.text == "Software"):
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Okay, I see that you are facing " + update.message.text +
                                        " issues. Click on the option which best applies to your situation.",
                                        reply_markup=app_issues_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_app_issues_keyboard).strip('"') + "\n\n")

        log_file.close()

        #setting value for sql db
        sql_category = message.text

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return appIssueChoice


    elif (update.message.text == "System"):
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Okay, I see that you are facing " + update.message.text +
                                        " issues. Click on the option which best applies to your situation.",
                                        reply_markup=system_issues_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_system_issues_keyboard).strip('"') + "\n\n")

        log_file.close()

        #setting value for sql db
        sql_category = message.text

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return systemIssueChoice


    elif (update.message.text == "User Enquiry"):
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Okay, I see that you are facing " + update.message.text +
                                        " issues. Click on the option which best applies to your situation.",
                                        reply_markup=userenquiry_issues_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_userenquiry_issues_keyboard).strip('"') + "\n\n")

        log_file.close()

        #setting value for sql db
        sql_category = message.text

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return userenquiryIssueChoice


    else:
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip('"') + "\n\n")

        reply = await update.message.reply_text("Sorry, but I will need to get my Customer Care colleague to help you with this. \n\n"
                                        "Please dial +65 12345678 to reach my colleague.",
                                        reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))   


async def tablet_issue_choice(update, context):
    #bot details log
    global current_datetime
    global sql_issue_reported
    global user_question

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    if (update.message.text == "Tablet cannot turn on"):
        #restart timer
        context.user_data['timer_task'].cancel()

        user_question = message.text

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip('"') + "\n\n")

        reply = await update.message.reply_text("Here are some methods you may try to solve your probem: \n"
                                        "1. Press and hold the Power button for at least 5 seconds. \n"
                                        "2. Charge the tablet: Connect the tablet's charging wire to the tablet and to a power outlet.",
                                        reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip('"') + "\n\n")
        
        time.sleep(10)

        #setting value for sql db
        sql_issue_reported = message.text

        solve_issue_keyboard = [["Yes"], ["No"]]

        #keyboard formatting
        formatted_list = [f"['{item}']" for inner_list in solve_issue_keyboard for item in inner_list]
        new_solve_issue_keyboard = ', '.join(formatted_list)

        solve_issue_reply_markup = ReplyKeyboardMarkup(solve_issue_keyboard, resize_keyboard=True, one_time_keyboard=True)
        reply = await update.message.reply_text("Were you able to solve your problem?",
                                        reply_markup=solve_issue_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")
        log_file.write("Keyboard Options: " + repr(new_solve_issue_keyboard).strip('"') + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return solveProblem


    elif (update.message.text == "Tablet is flickering"):
        #restart timer
        context.user_data['timer_task'].cancel()

        user_question = message.text

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("I will guide you through a soft reset to attempt to fix the problem: \n\n"
                                        "1. Press and hold the Power button until a Tablet Option box pops up \n"
                                        "2. Select Restart \n"
                                        "3. Wait until the device turns off for a reboot \n"
                                        "4. Check if the screen flickering issue still occurs",
                                        reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")  
        
        time.sleep(10)

        #setting value for sql db
        sql_issue_reported = message.text

        solve_issue_keyboard = [["Yes"], ["No"]]

        #keyboard formatting
        formatted_list = [f"['{item}']" for inner_list in solve_issue_keyboard for item in inner_list]
        new_solve_issue_keyboard = ', '.join(formatted_list)

        solve_issue_reply_markup = ReplyKeyboardMarkup(solve_issue_keyboard, resize_keyboard=True, one_time_keyboard=True)
        reply = await update.message.reply_text("Were you able to solve your problem?",
                                        reply_markup=solve_issue_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + (repr(reply.text).strip("'")).replace(" \n", "") + "\n")
        log_file.write("Keyboard Options: " + repr(new_solve_issue_keyboard).strip('"') + "\n\n")

        log_file.close()


        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return solveProblem


    else:
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Please type in the question you would like to ask.",
                                reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))
        
        return newQuestion



async def app_issue_choice(update, context):
    #bot details log
    global current_datetime
    global sql_status
    global sql_issue_reported
    global user_chat_id 
    global user_question

    chat = update.message.chat
    user_chat_id = chat.id

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    if (update.message.text == "Application cannot be opened"):
        #restart timer
        context.user_data['timer_task'].cancel()

        user_question = message.text

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Here are some methods you may try to solve your probem: \n"
                                        "1. Restart the app. \n"
                                        "2. Restart the tablet: Press and hold the Power button and click on 'Restart'.",
                                        reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip('"') + "\n\n")
        
        time.sleep(10)

        #setting value for sql db
        sql_issue_reported = message.text

        solve_issue_keyboard = [["Yes"], ["No"]]

        #keyboard formatting
        formatted_list = [f"['{item}']" for inner_list in solve_issue_keyboard for item in inner_list]
        new_solve_issue_keyboard = ', '.join(formatted_list)

        solve_issue_reply_markup = ReplyKeyboardMarkup(solve_issue_keyboard, resize_keyboard=True, one_time_keyboard=True)
        reply = await update.message.reply_text("Were you able to solve your problem?",
                                        reply_markup=solve_issue_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_solve_issue_keyboard).strip('"') + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return solveProblem


    elif (update.message.text == "Application not working as expected"):
        #restart timer
        context.user_data['timer_task'].cancel()

        user_question = message.text

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")


        reply = await update.message.reply_text("I would need to get my Customer Care colleague to help you with this. \n"
                                        "He/She will get back to you shortly.",
                                        reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")

        log_file.close()

        #setting value for sql db
        sql_status = "Open"
        sql_issue_reported = message.text

        #send to admin chat
        admin_chat_id = -4064577558
        admin_bot_msg = "Chat ID: " + repr(sql_chat_id) + "\nUnit: " + repr(sql_unit).strip("'") + "\n\nCategory: " + repr(sql_category).strip("'") + "\nIssue Reported: " + sql_issue_reported.strip("'")
        to_url = 'https://api.telegram.org/bot{}/sendMessage?chat_id={}&text={}&parse_mode=HTML'.format(token, admin_chat_id, admin_bot_msg)
        requests.get(to_url)


    else:
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Please type in the question you would like to ask.",
                                reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))
        
        return newQuestion



async def system_issue_choice(update, context):
    #bot details log
    global current_datetime
    global sql_status
    global sql_issue_reported
    global user_chat_id
    global user_question

    chat = update.message.chat
    user_chat_id = chat.id

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    if (update.message.text == "How do I create a new Windows account?"):
        #restart timer
        context.user_data['timer_task'].cancel()

        user_question = message.text

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Follow the steps below: \n"
                                        "1. Open the Control Panel. \n"
                                        "2. Click on 'User Accounts'. \n"
                                        "3. Click 'Create a new account'.",
                                        reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip('"') + "\n\n")
        
        time.sleep(10)

        #setting value for sql db
        sql_issue_reported = message.text

        solve_issue_keyboard = [["Yes"], ["No"]]

        #keyboard formatting
        formatted_list = [f"['{item}']" for inner_list in solve_issue_keyboard for item in inner_list]
        new_solve_issue_keyboard = ', '.join(formatted_list)

        solve_issue_reply_markup = ReplyKeyboardMarkup(solve_issue_keyboard, resize_keyboard=True, one_time_keyboard=True)
        reply = await update.message.reply_text("Were you able to solve your problem?",
                                        reply_markup=solve_issue_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_solve_issue_keyboard).strip('"') + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return solveProblem


    else:
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Please type in the question you would like to ask.",
                                reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))
        
        return newQuestion



async def userenquiry_issue_choice(update, context):
    #bot details log
    global current_datetime
    global sql_status
    global sql_issue_reported
    global user_chat_id
    global user_question

    chat = update.message.chat
    user_chat_id = chat.id

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    if (update.message.text == "How to do monthly patching"):
        #restart timer
        context.user_data['timer_task'].cancel()

        user_question = message.text

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Patching guide: \n"
                                                "1. Download all the files from Google Drive. \n"
                                                "2. Copy it to the S3 Tablet. \n"
                                                "3. Login to system admin. \n",
                                                reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip('"') + "\n\n")
        
        time.sleep(10)

        #setting value for sql db
        sql_issue_reported = message.text

        solve_issue_keyboard = [["Yes"], ["No"]]

        #keyboard formatting
        formatted_list = [f"['{item}']" for inner_list in solve_issue_keyboard for item in inner_list]
        new_solve_issue_keyboard = ', '.join(formatted_list)

        solve_issue_reply_markup = ReplyKeyboardMarkup(solve_issue_keyboard, resize_keyboard=True, one_time_keyboard=True)
        reply = await update.message.reply_text("Were you able to solve your problem?",
                                        reply_markup=solve_issue_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_solve_issue_keyboard).strip('"') + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return solveProblem


    elif (update.message.text == "How to give approval rights"):
        #restart timer
        context.user_data['timer_task'].cancel()

        user_question = message.text

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Follow the steps below: \n"
                                                "1. Open the S3 application. \n"
                                                "2. In the S3 admin account > select 'Approver mgt' module. \n"
                                                "3. Key in the NRIC of the new approver.",
                                        reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip('"') + "\n\n")
        
        time.sleep(10)

        #setting value for sql db
        sql_issue_reported = message.text

        solve_issue_keyboard = [["Yes"], ["No"]]

        #keyboard formatting
        formatted_list = [f"['{item}']" for inner_list in solve_issue_keyboard for item in inner_list]
        new_solve_issue_keyboard = ', '.join(formatted_list)

        solve_issue_reply_markup = ReplyKeyboardMarkup(solve_issue_keyboard, resize_keyboard=True, one_time_keyboard=True)
        reply = await update.message.reply_text("Were you able to solve your problem?",
                                        reply_markup=solve_issue_reply_markup)

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n")
        log_file.write("Keyboard Options: " + repr(new_solve_issue_keyboard).strip('"') + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))


        return solveProblem


    else:
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Please type in the question you would like to ask.",
                                reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

        log_file.close()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))
        
        return newQuestion


async def new_question(update, context):
    #restart timer
    context.user_data['timer_task'].cancel()

    global current_datetime
    global sql_issue_reported
    global sql_status
    global user_chat_id
    global user_question

    chat = update.message.chat
    user_chat_id = chat.id

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    #bot details log
    log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")
    
    reply = await update.message.reply_text("Alright, my Customer Care colleague will get back to you shortly.",
                                            reply_markup=ReplyKeyboardRemove())

    #bot details log
    bot_current_date = datetime.now().strftime("%H:%M:%S")
    log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

    log_file.close()

    user_question = message.text

    #setting value for sql db
    sql_issue_reported = user_question
    sql_status = "Open"


    #send to admin chat
    admin_chat_id = -4064577558
    admin_bot_msg = "Chat ID: " + repr(sql_chat_id) + "\nUnit: " + repr(sql_unit).strip("'") + "\n\nCategory: " + repr(sql_category).strip("'") + "\nIssue Reported: " + sql_issue_reported.strip("'")
    to_url = 'https://api.telegram.org/bot{}/sendMessage?chat_id={}&text={}&parse_mode=HTML'.format(token, admin_chat_id, admin_bot_msg)
    requests.get(to_url)



async def solve_problem(update, context):
    #bot details log
    global current_datetime
    global sql_date_closed
    global sql_status
    global sql_response_time
    global user_chat_id

    chat = update.message.chat
    user_chat_id = chat.id

    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")

    if (update.message.text == "Yes"):
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("Wonderful! \n\n"
                                        "Is there anything else I can assist you with?",
                                        reply_markup=ReplyKeyboardRemove())

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

        log_file.close()

        #setting value for sql db
        sql_status = "Closed"
        sql_date_closed = datetime.now()
        sql_response_time = None

        insert_into_sqldb()

        #create timer using asyncio create task
        context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))



    elif (update.message.text == "No"):
        #restart timer
        context.user_data['timer_task'].cancel()

        #bot details log
        log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

        reply = await update.message.reply_text("I would need to get my Customer Care colleague to help you with this. \n"
                                        "He/She will get back to you shortly.",
                                        reply_markup=ReplyKeyboardRemove()) 

        #bot details log
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

        log_file.close()


        #setting value for sql db
        sql_status = "Open"


        #send to admin chat
        admin_chat_id = -4064577558
        admin_bot_msg = "Chat ID: " + repr(sql_chat_id) + "\nUnit: " + repr(sql_unit).strip("'") + "\n\nCategory: " + repr(sql_category).strip("'") + "\nIssue Reported: " + sql_issue_reported.strip("'") + " - Solution provided did not solve problem"
        to_url = 'https://api.telegram.org/bot{}/sendMessage?chat_id={}&text={}&parse_mode=HTML'.format(token, admin_chat_id, admin_bot_msg)
        requests.get(to_url)



async def bot_help(update, context):
    #restart timer
    context.user_data['timer_task'].cancel()

    #bot details log
    global current_datetime
    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a", encoding='utf-8')

    log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

    reply = await update.message.reply_text("I am a bot, designed to help answer any queries you may have. \n\n"
                                    "I can respond to commands shown in the menu, as well as keyboard options you pick each time I prompt you. \n\n"
                                    "I will always do my best to help you 💪 , but if I am unable to solve your problem, I will direct you to my Customer Care colleague. \n\n\n"
                                    "So, feel free to ask away! 😊 ",
                                    reply_markup=ReplyKeyboardRemove())

    #bot details log
    bot_current_date = datetime.now().strftime("%H:%M:%S")
    log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

    log_file.close()

    #create timer using asyncio create task
    context.user_data['timer_task'] = asyncio.create_task(wait_for_timeout(update, context))



async def cancel(update, context):
    #restart timer
    context.user_data['timer_task'].cancel()

    #bot details log
    global current_datetime
    message = update.message
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a", encoding='utf-8')

    log_file.write("[" + repr((message.date + timedelta(hours=8)).strftime("%H:%M:%S")).strip("'") + "] " + chat_name + ": " + repr(message.text).strip("'") + "\n\n")

    reply = await update.message.reply_text("Thank you for using S3 Bot, I hope to have effectively assisted you today! \n\n"
                                    "If you would like to start a conversation with me again, simply enter /start, or head over to the menu tab for a list of commands, and I will be there to assist you! \n\n"
                                    "Bye! 👋",
                                    reply_markup=ReplyKeyboardRemove())

    #bot details log
    bot_current_date = datetime.now().strftime("%H:%M:%S")
    log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

    log_file.close()

    return ConversationHandler.END


async def wait_for_timeout(update, context):
    try:
        await asyncio.sleep(300)
        await bot_timeout(update, context)
    except asyncio.CancelledError:
        pass


async def bot_timeout(update, context):
    reply = await update.message.reply_text("Seems like you do not need any more help for now, and so I will end our conversation first! \n\n"
                                    "But if you need me again, click /start, or head over to the menu tab for a list of commands, and I will be there to assist you! \n\n"
                                    "Thank you for using S3 Bot, and I hope to have effectively assisted you today! \n\n\n"
                                    "Bye! 👋",
                                    reply_markup=ReplyKeyboardRemove())

    #bot details log
    global current_datetime
    log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a", encoding='utf-8')

    bot_current_date = datetime.now().strftime("%H:%M:%S")
    log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(reply.text.replace("\n", "")).strip("'") + "\n\n")

    log_file.close()


def admin_reply(update, context):
    global user_chat_id 
    global user_question
    global sql_response_time
    global chat_name
    global current_datetime

    message = update.message
    msg_chat_id = message.chat_id

    
    #send msg back to user 
    if (msg_chat_id == -4064577558):
        admin_reply_msg = "Question: \n" + user_question + " \n\n" + "Answer: \n" + message.text
        to_url = 'https://api.telegram.org/bot{}/sendMessage?chat_id={}&text={}&parse_mode=HTML'.format(token, user_chat_id, admin_reply_msg)
        requests.get(to_url)

        #setting value for sql db
        sql_response_time = datetime.now()

        #bot details log
        log_file = open(f"{file_dir}\\{chat_name}_{current_datetime}.log", "a")
    
        bot_current_date = datetime.now().strftime("%H:%M:%S")
        log_file.write("[" + bot_current_date + "] " + "Bot: " + repr(admin_reply_msg.replace("\n", "")).strip("'") + "\n\n")

        log_file.close()

        insert_into_sqldb()


def insert_into_sqldb():
    #writing data to db
    sql_id = uuid.uuid4()

    if (sql_id!= None) and (sql_chat_id != None) and (sql_date != None) and (sql_time != None) and (sql_unit != None) and (sql_reported_by != None) and (sql_category != None) and (sql_issue_reported != None) and (sql_status != None):
        cursor.execute('''
        INSERT INTO bot_details 
            ([id],
            [chat_id],
            [date],
            [time],
            [unit],
            [reported_by],
            [category],
            [issue_reported],
            [response_time],
            [date_closed],
            [status],
            [remark])
        VALUES
            (?, 
            ?,
            CONVERT(date, ?),
            CONVERT(time, ?),
            ?,
            ?,
            ?,
            ?,
            CONVERT(datetime, ?),
            CONVERT(date, ?),
            ?,
            ?)
        ''', (sql_id, sql_chat_id, sql_date, sql_time, sql_unit, sql_reported_by.strip("'"), sql_category.strip("'"), sql_issue_reported.strip('"').strip("'"), sql_response_time, sql_date_closed, sql_status, sql_remarks))

        connection.commit()


    else:
        pass



async def generate_excel(update, context):
    #writing data to excel
        excel_id_list = []
        excel_chat_id_list = []
        excel_date_list = []
        excel_time_list = []
        excel_unit_list = []
        excel_reported_by_list = []
        excel_category_list = []
        excel_issue_reported_list = []
        excel_response_time_list = []
        excel_date_closed_list = []
        excel_status_list = []
        excel_remark_list = []
 

        cursor.execute('''
              SELECT [id],
              [chat_id],
              [date],
              [time],
              [unit],
              [reported_by],
              [category],
              [issue_reported],
              [response_time],
              [date_closed],
              [status],
              [remark]
              FROM [telegrambot].[dbo].[bot_details]
            ''')

        retrieve_for_excel = cursor.fetchall()
        for i in retrieve_for_excel:
            excel_id_list.append(i[0])
            excel_chat_id_list.append(i[1])
            excel_date_list.append(i[2])
            excel_time_list.append(i[3])
            excel_unit_list.append(i[4])
            excel_reported_by_list.append(i[5])
            excel_category_list.append(i[6])
            excel_issue_reported_list.append(i[7])
            if i[8] == None:
                excel_response_time_list.append("NA")
            else:
                excel_response_time_list.append(i[8])
            excel_date_closed_list.append(i[9])
            excel_status_list.append(i[10])
            excel_remark_list.append(i[11])

        output_df = pd.DataFrame({"ID": excel_id_list, "Chat ID": excel_chat_id_list,"Date": excel_date_list,
                    "Time": excel_time_list, "Unit": excel_unit_list, "Reported By": excel_reported_by_list, 
                    "Category": excel_category_list, "Issue Reported": excel_issue_reported_list, "Response Time": excel_response_time_list, 
                    "Date Closed": excel_date_closed_list, "Status": excel_status_list, "Remarks": excel_remark_list})

        excel_datetime = datetime.now().strftime('%d %B %Y_%H.%M.%S')
        excel_file_path = data["excel_file_path"].format(conf_excel_datetime=excel_datetime)

        output_df.to_excel(excel_file_path, index=False)


        #formatting of cells
        active_workbook = openpyxl.load_workbook(excel_file_path)
        active_worksheet = active_workbook.active

        #setting column widths
        column_widths = {'A': 40, 'B': 15, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 15, 'H': 25, 'I': 20, 'J': 15, 'K': 15, 'L': 25}
        for column, width in column_widths.items():
            active_worksheet.column_dimensions[column].width = width
        
        #setting column style
        range = active_worksheet['A':'L']
        for cell in range:
            for c in cell:
                c.alignment = Alignment(horizontal='left', vertical='top', wrapText=True)

        active_workbook.save(excel_file_path)

        #adding filter
        active_worksheet_for_filter = active_workbook['Sheet1']
        filter_range = 'B:K'
        active_worksheet_for_filter.auto_filter.ref = filter_range
        active_workbook.save(excel_file_path)

        await update.message.reply_text("Excel successfully generated. \n\n"
                                        "File Path: " + excel_file_path,
                                    reply_markup=ReplyKeyboardRemove())



application = Application.builder().token("6364510958:AAHghv-nCZ8R2eZcivuiFNNFZU-IQZ2Boas").build()

conversation_handler = ConversationHandler(
    entry_points = [CommandHandler('start', start)],
    states = {
        editName: [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_name)],
        editUnit: [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_unit)],
        startUnit: [MessageHandler(filters.TEXT & ~filters.COMMAND, start_unit)],
        startName: [MessageHandler(filters.TEXT & ~filters.COMMAND, start_name)],
        issueChoice: [MessageHandler(filters.TEXT & ~filters.COMMAND, issue_choice)],
        tabletIssueChoice: [MessageHandler(filters.TEXT & ~filters.COMMAND, tablet_issue_choice)],
        appIssueChoice: [MessageHandler(filters.TEXT & ~filters.COMMAND, app_issue_choice)],
        systemIssueChoice: [MessageHandler(filters.TEXT & ~filters.COMMAND, system_issue_choice)],
        userenquiryIssueChoice: [MessageHandler(filters.TEXT & ~filters.COMMAND, userenquiry_issue_choice)],
        newQuestion: [MessageHandler(filters.TEXT & ~filters.COMMAND, new_question)],
        solveProblem: [MessageHandler(filters.TEXT & ~filters.COMMAND, solve_problem)]
    },
    fallbacks = [CommandHandler('cancel', cancel)],
    conversation_timeout=300,
    allow_reentry=True
)

application.add_handler(conversation_handler)

application.add_handler(CommandHandler('help', bot_help))

admin_message_handler = MessageHandler(filters.TEXT & ~filters.COMMAND, admin_reply)
application.add_handler(admin_message_handler)

application.add_handler(CommandHandler('generate', generate_excel))

#application.run_polling(allowed_updates=Update.ALL_TYPES)

