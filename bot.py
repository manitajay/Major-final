import os

import telebot

import attendance as at

# BOT_TOKEN = os.environ.get('BOT_TOKEN')
BOT_TOKEN='5960792239:AAH3BVTxIBOwfK-6pU_p_8YzhiS-eHldZrQ'

bot = telebot.TeleBot(BOT_TOKEN)

print(BOT_TOKEN)


print(at.data[0][0])


str="manit"
while(1):
    # print("tu")
    @bot.message_handler(func=lambda msg: True)
    def fun(message):
        str=message.text
        bot.reply_to(message,"hii")


        if(str=="exit"):
            bot.reply_to(message,"hey")








# tmp="node"
# @bot.message_handler(commands=['start', 'hello'])
# def send_welcome(message):
#     bot.reply_to(message, "Hello manitian , please enter , <SCHOLAR_NO/Subject Code> <Password>:")
#     tmp=message.text
#     login=True




# @bot.message_handler(func=lambda msg: True)
# def echo_all(message):
#     print(message.text)
#     str=message.text

    # if(str == "login"):
    #     bot.reply_to(message,"for login as student reply 1 or login as professor reply 2")
    #     @bot.message_handler(func=lambda msg:True)
    #     def fun(message):
    #         str=message.text

    #         if(str=="1"):
    #             bot.reply_to(message, "Hello manitian , please enter , <SCHOLAR_NO><Password>(without space):")
    #         elif(message.text=="2"):
    #             bot.reply_to(message,"Hello sir/mam , please enter , <SUBJECT CODE><Password>(withoutspace):")
    #         else:
    #             bot.reply_to(message,"wrong input")

    # bot.reply_to(message, message.text)


# print(tmp)

# @bot.message_handler(commands=['id'])
# def send_welcome(message):
#     bot.reply_to(message, tmp)


    

#     login=True






# if(login):



# @bot.message_handler(commands=['TEMP'])
# def send_welcome(message):
#     bot.reply_to(message,"manit temprature is __ degree celcius")


# @bot.message_handler(commands=[])


bot.infinity_polling()

