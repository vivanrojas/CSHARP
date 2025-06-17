using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

namespace EndesaBusiness.utilidades
{
    public class TelegramMensajes
    {

        string token;
        string channel;

        public TelegramMensajes(string telegram_token, string telegram_channel)
        {
            token = telegram_token;
            channel = telegram_channel;
        }

        public void SendMessage(string text)
        {
            // string token = "1573051703:AAHpt7_UOGlX3pHzg4y6ZOHMJhAySaC2sT8"; // Token Contratacion
            // string channel = "-1001466114152"; // Contratacion
            //string text = "Esto es un mensaje de prueba de Gabriel"; 
          

            string url = "";

            WebClient client = new WebClient();
            try
            {
                
                client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.33 Safari/537.36");
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                url = @"https://api.telegram.org/bot" + token + "/sendMessage?chat_id=" + channel+ "&text=" + text;

                client.UploadString(url,"");


                //using (var httpClient = new HttpClient())
                //{
                //    var res = httpClient.GetAsync(
                //        $"https://api.telegram.org/bot{token}/sendMessage?chat_id={channel}&text={text}"
                //        ).Result;
                //    if (res.StatusCode == HttpStatusCode.OK)
                //    { /* done, go check your channel */ }
                //    else
                //        Console.WriteLine("Error SendMessage");

                //}
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }


        }



       
        //public void SendMessage(string text)
        //{
        //    // string token = "1573051703:AAHpt7_UOGlX3pHzg4y6ZOHMJhAySaC2sT8"; // Token Contratacion
        //    // string channel = "-1001466114152"; // Contratacion
        //    //string text = "Esto es un mensaje de prueba de Gabriel"; 
        //    try
        //    {



        //        using (var httpClient = new HttpClient())
        //        {
        //            var res = httpClient.GetAsync(
        //                $"https://api.telegram.org/bot{token}/sendMessage?chat_id={channel}&text={text}"
        //                ).Result;
        //            if (res.StatusCode == HttpStatusCode.OK)
        //            { /* done, go check your channel */ }
        //            else
        //             Console.WriteLine("Error SendMessage"); 
                    
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //    }


        //}


        //        public async Task<Message> SendMessage(string text)
        //        {

        //            var botClient = new Telegram.Bot.Api(token);
        //            return  botClient.SendTextMessage("@PruebasEndesa_Bot", text);

        //);
        //        }



        //public async Task SendMessage(string texto)
        //{
        //    try
        //    {

        //        var botClient = new TelegramBotClient(token);

        //        Message message = await botClient.SendTextMessageAsync(
        //            chatId: channel, // or a chat id: 123456789
        //            text: texto,
        //            parseMode: ParseMode.Markdown,
        //        disableNotification: false,
        //        replyMarkup: new InlineKeyboardMarkup(InlineKeyboardButton.WithUrl(
        //        "Check sendMessage method",
        //        "https://core.telegram.org/bots/api#sendmessage"
        //                )));
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //    }
        //}

    }
}
